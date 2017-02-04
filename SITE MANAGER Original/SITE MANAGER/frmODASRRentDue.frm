VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRViewer.dll"
Begin VB.Form frmODASRRentDue 
   Caption         =   "Rent Payments"
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
Attribute VB_Name = "frmODASRRentDue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As CRAXDRT.Report

Private Sub Form_Activate()
If bexportRECORD = True Then Unload Me

End Sub

Private Sub Form_Load()
Dim f, G As String
On Error GoTo errMSG

    G = Format(frmODASSearchSitesNotPaid.txtLastDate.Value, "dd/MM/yyyy")
If frmODASSearchSitesNotPaid.txtStartDate.Visible = True Then
    G = Format(frmODASSearchSitesNotPaid.txtStartDate.Value, "dd/MM/yyyy") & " AND " & Format(frmODASSearchSitesNotPaid.txtLastDate.Value, "dd/MM/yyyy")
End If

With frmODASSearchSitesNotPaid
    Select Case .strReport
        Case "PendingPayment"
            Set Report = New ODASRRentDue
        Case "PendingPaymentAsAtASingleDate"
            Set Report = New ODASRRentDue
        Case "VouchersPrepared"
            Set Report = New ODASRRentDue
        Case "PendingConfirmation"
            Set Report = New ODASRRentDue
        Case "PaymentsConfirmed"
             Set Report = New ODASRRentPaid
        Case Else
            Set Report = New ODASRRentDue
    End Select
End With


Screen.MousePointer = vbHourglass: Dim myCrystalLogon As New clsCrystallogon
CRViewer1.ReportSource = Report: Report.Database.LogOnServer "pdsodbc.dll", DSN, Database, Uid, Pwd: CRViewer1.ReportSource = Report: DoEvents
Set myCrystalLogon.CrxRep = Report: myCrystalLogon.setCRLoginInfo:

Report.DiscardSavedData

Report.SQLQueryString = strSQL
Report.Text13.SetText (frmODASSearchSitesNotPaid.Caption)
Report.Text15.SetText G

If bexportRECORD = True Then
    Report.ExportOptions.DestinationType = crEDTDiskFile
    Report.ExportOptions.DiskFileName = App.Path & "\SitePaymentStatus.xls"
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
