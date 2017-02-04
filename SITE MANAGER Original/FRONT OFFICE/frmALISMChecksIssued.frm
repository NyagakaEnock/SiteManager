VERSION 5.00
Begin VB.Form frmALISMCheckIssued 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque Issued"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmALISMCheckIssued"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCOLLECTED As clsALISCheque

Private Sub cbobankNo_GotFocus()
        strSQL = "SELECT * FROM ALISPBankAccount"
        bankNoGotFocus
End Sub

Private Sub cboBankNo_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cbobankNo_LostFocus()
        strSQL = "SELECT * FROM ALISPBankAccount WHERE Details = '" & cbobankNo.Text & "'"
        BankNoLostFocus
End Sub

Private Sub cboIdType_GotFocus()
    SelectIDTypeGotFocus
End Sub

Private Sub cboIdType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboIdType_LostFocus()
    SelectIDTypeLostFocus
End Sub

Private Sub cmdAddNew_Click()
        Set rsCOLLECTED = New clsALISCheque
        rsCOLLECTED.clearCHEQUE
        Set rsCOLLECTED = Nothing

        enableALLRECORD
        disableButtons
End Sub


Private Sub cmdCancel_Click()
        Set rsCOLLECTED = New clsALISCheque
        rsCOLLECTED.Cancelrecord
        Set rsCOLLECTED = Nothing
        bmakePAYMENT = False
        breversePAYMENT = False
        bissueCHECKS = False

End Sub



Private Sub cmdUpdate_Click()
        Set rsCOLLECTED = New clsALISCheque
        bsaveRECORD = False
        rsCOLLECTED.UpdateChecksIssued
        Set rsCOLLECTED = Nothing

End Sub

Private Sub DTPickerChequeDate_Change()
        Set rsCOLLECTED = New clsALISCheque
        rsCOLLECTED.ChangeDATE
        Set rsCOLLECTED = Nothing

End Sub

Private Sub Form_Activate()
        bmakePAYMENT = False
        breverpayment = False
        bissueCHECKS = True
        
        Set rsCOLLECTED = New clsALISCheque
        If bissueCHECKS = True Then
        
            strSQL = "Select * from ALISMCheque where chequeNo = '" & frmALISMCheckIssued.txtChequeNo & "'"
        End If
        rsCOLLECTED.loadAPPROVEDCHECKS
        Set rsCOLLECTED = Nothing
        disableALLRECORD
        enableButtons
End Sub

Private Sub Form_Load()
        'OpenConnection
End Sub


Private Sub Form_Unload(Cancel As Integer)
        bmakePAYMENT = False
        breversePAYMENT = False
End Sub

Private Sub txtAmountPaid_LostFocus()
        Set rsCOLLECTED = New clsALISCheque
        rsCOLLECTED.checkSTATUS
        Set rsCOLLECTED = Nothing
End Sub
