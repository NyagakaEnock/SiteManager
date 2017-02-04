VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRViewer.dll"
Begin VB.Form frmRptODASSitesToExpire 
   Caption         =   "Form1"
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmRptODASSitesToExpire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New RptODASSitesToExpire
Public strReport As String

Private Sub Form_Activate()
If bexportRECORD = True Then Unload Me

End Sub

Private Sub Form_Load()
Dim f, G As String
On Error GoTo errMSG
Screen.MousePointer = vbHourglass: Dim myCrystalLogon As New clsCrystallogon
Set myCrystalLogon.CrxRep = Report: myCrystalLogon.setCRLoginInfo:
CRViewer1.ReportSource = Report
Report.SQLQueryString = strSQL
If frmRptODASSitesToExpire.strReport = "" Then
        'Report.txtStartDate.SetText frmODASSearchSiteNewSites.txtStartDate.Text
        'Report.txtExpiryDate.SetText frmODASSearchSiteNewSites.txtLastDate.Text
        Report.txtReportTitle.SetText "Sites expiring within the period Starting from " & Format(frmODASSitesToExpire.txtStartDate.Text, "dd/MM/yyyy") & " to " & Format(frmODASSitesToExpire.txtLastDate.Text, "dd/MM/yyyy") & ";"
Else
        'Report.txtExpiryDate.SetText frmODASSearchSiteNewSites.txtLastDate.Text
        Report.txtReportTitle.SetText "Sites to expire by " & Format(frmODASSitesToExpire.txtLastDate.Text, "dd/MM/yyyy")

End If
Report.DiscardSavedData
If bexportRECORD = True Then
    Report.ExportOptions.DestinationType = crEDTDiskFile
    Report.ExportOptions.DiskFileName = App.Path & "\ExpiringSites.xls"
    Report.ExportOptions.FormatType = crEFTExcel50
    Report.Export False
    
    Dim oXLApp As Object, oXLWorkbook As Object
    
    Set oXLApp = CreateObject("Excel.Application")
    
    Set oXLWorkbook = oXLApp.Workbooks.Open(Report.ExportOptions.DiskFileName)
    
    oXLApp.Visible = True
    
Else
    CRViewer1.ViewReport
End If
Screen.MousePointer = vbDefault
Exit Sub
errMSG:
    ErrorMessage
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
