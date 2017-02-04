VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmLedgerDetails 
   Caption         =   "Ledger Details Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmLedgerDetails.frx":0000
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
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frmLedgerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Report As New rptledgerdetails
Dim rsAgent As ADODB.Recordset

Private Sub Form_Load()

On Error GoTo err

Screen.MousePointer = vbHourglass

        Report.RecordSelectionFormula = "{ODASMJobBrief.JobBriefNo}='" & frmALISMLedgerDetails.cboJobBriefNo & " ';"
        Report.txtpreparedby.SetText (UCase((CurrentUserName)))
        Report.Subreport1.OpenSubreport.RecordSelectionFormula = "{ODASMJobBriefLedger.DocumentNo}='" & frmALISMLedgerDetails.cboJobBriefNo & "';"

        Set rsAgent = New ADODB.Recordset
        rsAgent.Open "Select * from ODASMJobBrief,ALISPAgent where ODASMJobBrief.JobBriefNo='" & frmALISMLedgerDetails.cboJobBriefNo.Text & " ' and ODASMJobBrief.AgentNo=ALISPAgent.AgentNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
        Dim names As String
        
        If Not rsAgent.BOF And Not rsAgent.EOF Then
        
            names = rsAgent!CompanyName & rsAgent!OtherNames & rsAgent!titlecode
        End If
        
        Report.txtagentNames.SetText (names)
        CRViewer1.ReportSource = Report
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        Policy = False
        ledger = False
        Exit Sub
err:
        Screen.MousePointer = vbDefault

ErrorMessage
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
