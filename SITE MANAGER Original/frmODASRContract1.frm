VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmODASRContract1 
   Caption         =   "Print Contract"
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
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frmODASRContract1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New rptODASLandLordContract1

Private Sub Form_Load()
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "Select * From ODASMLeaseAgreement Where ContractNo = '" & CurrentRecord & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
        If rsFindRecord!Authorized = "Y" Then
            CRViewer1.EnablePrintButton = True
        Else
            CRViewer1.EnablePrintButton = False
        End If
        
    Set rsFindRecord = Nothing
    Screen.MousePointer = vbHourglass: Dim myCrystalLogon As New clsCrystallogon
    Report.RecordSelectionFormula = "{ODASMLeaseAgreement.ContractNo } = '" & CurrentRecord & "'"
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
