VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub updateRECORD()
On Error GoTo err
       With frmODASMInventory
            
                strSQL = "select * from ODASMInventory Where ProductCode = '" & .txtProductCode & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!ProductCode = .txtProductCode.Text
                        rsSAVE!PreparedBy = CurrentUserName
                        rsSAVE!DatePrepared = Date
                        rsSAVE!Prepared = "Y"
                End If
                
                rsSAVE!StockOnHand = .txtStockOnHand.Text
                rsSAVE!StockOnOrder = .txtStockOnOrder.Text
                rsSAVE!StockReceived = .txtStockReceived.Text
                rsSAVE!StockReserved = .txtStockReserved.Text
                rsSAVE!UnitCode = .txtUnitCode.Text
                rsSAVE!ReorderLevel = .txtReorderLevel.Text
                rsSAVE!ReorderStatus = .txtReOrderStatus.Text
                rsSAVE!LastUpdate = .txtLastUpDate.Text
                rsSAVE!LastLPONo = .txtLastLPONo.Text
                rsSAVE!LastLPODate = .txtLastLPODate.Text
                rsSAVE!LastLPOStatus = .txtLastLPOStatus.Text
                rsSAVE.Update
                rsSAVE.Requery
                bSaveRECORD = False
                Set rsSAVE = Nothing
                strSQL = Empty
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub SaveRECORD()
    ValidateRECORD
    If bSaveRECORD = True Then
            updateRECORD
    End If
End Sub

Private Sub ValidateRECORD()
On Error GoTo err
        With frmODASMInventory
            bSaveRECORD = False
            
            If .txtProductCode.Text = Empty Then
                    MsgBox "The Product Code enetered is Invalid"
                    .txtProductCode.SetFocus
            
            ElseIf .txtProductDescription.Text = Empty Then
                    MsgBox "The Product Description is Required"
                    .txtProductDescription.SetFocus
            
            ElseIf CDbl(.txtStockIssued) < 0 Then
                    MsgBox "The Stock Issued cannot be Zero"
                    .txtStockIssued.SetFocus
            
            ElseIf CDbl(.txtStockOnHand) < 0 Then
                    MsgBox "The Stock on Hand Cannot be Zero"
                    .txtStockOnHand.SetFocus
            
            ElseIf CDbl(.txtStockOnOrder) < 0 Then
                    MsgBox "The Stock on Order Cannot be Less Than Zero"
                    .txtStockOnOrder.SetFocus
            
            ElseIf CDbl(.txtStockReceived) < 0 Then
                    MsgBox "The Stock Received Cannot be Zero"
                    .txtStockReceived.SetFocus
            
            ElseIf CDbl(.txtStockReserved) < 0 Then
                    MsgBox "The Stock Reserved Cannot be less than Zero"
                    .txtStockReserved.SetFocus
                    
            ElseIf CDbl(.txtReorderLevel) < 0 Then
                    MsgBox "The Reorder Level Cannot be Less than Zero"
                    .txtReorderLevel.SetFocus
            Else
                    bSaveRECORD = True
            End If
            
        End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub loadRECORD()
On Error GoTo err
       With frmODASMInventory
            
                strSQL = "select * from ODASMInventory Where ProductCode = '" & .txtProductCode & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                .txtStockOnHand.Text = rsSAVE!StockOnHand & ""
                .txtStockOnOrder.Text = rsSAVE!StockOnOrder & ""
                .txtStockReceived.Text = rsSAVE!StockReceived & ""
                .txtStockReserved.Text = rsSAVE!StockReserved & ""
                .txtUnitCode.Text = rsSAVE!UnitCode & ""
                .txtReorderLevel.Text = rsSAVE!ReorderLevel & ""
                .txtReOrderStatus.Text = rsSAVE!ReorderStatus & ""
                .txtLastUpDate.Text = rsSAVE!LastUpdate & ""
                .txtLastLPONo.Text = rsSAVE!LastLPONo & ""
                .txtLastLPODate.Text = rsSAVE!LastLPODate & ""
                .txtLastLPOStatus.Text = rsSAVE!LastLPOStatus & ""
        End With
Exit Sub

err:
    ErrorMessage
End Sub

