VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASReceiverOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub LoadNewRECORD()
On Error GoTo err

    With frmODASMReceiveOrder
            .txtInvoiceDate.Text = Date
            .DTPickerInvoiceDate.Value = Date
            .txtDeliveryDate.Text = Date
            .DTPickerDeliveryDate.Value = Date
            .txtinvoiceNo.Text = Empty
            .txtDeliveryNoteNo.Text = Empty
            .txtGRNNo.Text = Empty
            .txtGRNDate.Text = Date
            .DTPickerGRNDate.Value = Date
    End With

Exit Sub

err:
    ErrorMessage
End Sub
