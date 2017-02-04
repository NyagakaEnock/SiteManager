VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmNewReceipt 
   Caption         =   "RECEIPT"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmNewReceipt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
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
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmNewReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New rptNewReceipt
Option Explicit

Private Sub Form_Load()

On Error GoTo err

        Screen.MousePointer = vbHourglass
        Dim rsbank As ADODB.Recordset
        Set rsbank = New ADODB.Recordset
        Dim translater As New cMoneyConverter
        Set translater = New cMoneyConverter
       
        figures = CDbl(frmODASMReceipt.txtReceiptAmount.Text)
        translater.CallConverter
       
        rsbank.Open "Select * from ALISPBank,ALISMReceiptNew where ALISMReceiptNew.ReceiptNo='" & frmODASMReceipt.txtReceiptNo & "' and ALISPBank.BankNo=ALISMReceiptNew.BankNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If Not rsbank.BOF And Not rsbank.EOF Then
                Report.txtbank.SetText (rsbank!CompanyName)
        End If
        
        With Report
                .RecordSelectionFormula = "{ALISMReceiptDetails.ReceiptNo}='" & frmODASMReceipt.txtReceiptNo.Text & "';"
                .txtAmount.SetText (inwords)
        End With
            
            CRViewer1.ReportSource = Report
            CRViewer1.ViewReport
            Screen.MousePointer = vbDefault
    Exit Sub
err:             ErrorMessage
End Sub

Private Sub Form_Resize()
        CRViewer1.Top = 0
        CRViewer1.Left = 0
        CRViewer1.Height = ScaleHeight
        CRViewer1.Width = ScaleWidth
End Sub

