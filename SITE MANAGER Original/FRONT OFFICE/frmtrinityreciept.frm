VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmtrinityreciept 
   Caption         =   "Receipt Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   Icon            =   "frmtrinityreciept.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6330
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
Attribute VB_Name = "frmtrinityreciept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Report As New rpttrinityreciept

Option Explicit

Private Sub Form_Load()

'On Error GoTo err
        Screen.MousePointer = vbHourglass
        Dim banktype As String
        Dim rsbanktype As New adodb.Recordset
        Set rsbanktype = New adodb.Recordset
        Dim translater As New cMoneyConverter
        Set translater = New cMoneyConverter

       rsbanktype.Open "Select ALISPBank.LocalBank from ALISMReceipt,ALISPBank where ALISMReceipt.BankNo=ALISPBank.BankNo and ALISMReceipt.txtReceiptNo= '" & Screen.ActiveForm.txtReceiptNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
       
       figures = Screen.ActiveForm.txtReceiptAmount.Text
       translater.CallConverter

        With Report
                .RecordSelectionFormula = "{ALISMReceipt.txtReceiptNo}=" & CLng(Screen.ActiveForm.txtReceiptNo.Text) & ";"
                .cname.SetText ("Trinity Life Assurance Company Limited")
                .cpostaladdress.SetText ("P.O Box 12043-00400 Nairobi.")
                .ctel.SetText ("Tel: " & "244282/244229" & ".")
                .rpttxtamount.SetText (inwords)
                .upcountry.SetText ("")
                .txtLocal.SetText ("")
            If Screen.ActiveForm.txtChequeNo = "" Then GoTo continue
                
                If rsbanktype!LocalBank = "none" Then
                    .txtLocal.SetText ("")
                    '.txtReceiptAmount.SetText ("")
                    .upcountry.SetText (figures)
                Else
                    .upcountry.SetText ("")
                    '.txtReceiptAmount.SetText ("")
                    .txtLocal.SetText (figures)
                End If
continue:
                If Screen.ActiveForm.txtChequeNo = "" Then
                    .upcountry.SetText ("")
                    .txtLocal.SetText ("")
                    '.txtReceiptAmount.SetText (figures)
                End If
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
