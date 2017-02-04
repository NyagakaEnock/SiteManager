VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASReceiveOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub LoadNewRECORD()
On Error GoTo err

    With Screen.ActiveForm
            .txtInvoiceDate.Text = Empty
            .DTPickerInvoiceDate.Value = Date
            .txtDeliveryDate.Text = Empty
            .DTPickerDeliveryDate.Value = Date
            .txtInvoiceNo.Text = Empty
            .txtDeliveryNoteNo.Text = Empty
            .txtGRNNo.Text = Empty
            .txtGRNDate.Text = Date
            .DTPickerGRNDate.Value = Date
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadDELIVERYDETAILS()
On Error GoTo err

    With frmODASMProductSetup
            
            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASMGRN WHERE ODASMGRN.GRNNo = '" & .txtGRNNo.Text & "' ;"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            If rsCONTROL.BOF Or rsCONTROL.EOF Then Exit Sub
            
            .txtInvoiceDate.Text = rsCONTROL!InvoiceDate & ""
            .txtDeliveryDate.Text = rsCONTROL!DeliveryDate & ""
            .txtInvoiceNo.Text = rsCONTROL!InvoiceNo & ""
            .txtDeliveryNoteNo.Text = rsCONTROL!DeliveryNoteNo & ""
            .txtGRNNo.Text = rsCONTROL!GRNNo
            .txtGRNDate.Text = rsCONTROL!GrnDate & ""
    End With

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub loadPuchaseOrder()
On Error GoTo err
        With Screen.ActiveForm

            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASMPurchaseOrder WHERE OrderNo = '" & .txtOrderNo.Text & "' ;"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
            If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub

            .txtOrderDate.Text = rsCONTROL!OrderDate
            .txtOrderDescription.Text = rsCONTROL!OrderDescription
            .txtSupplierCode.Text = rsCONTROL!supplierCode
            .txtRemarks.Text = rsCONTROL!Remarks & ""
            .txtDeadlineDate.Text = rsCONTROL!DateRequired
            .txtTotalCost.Text = rsCONTROL!TotalCost
            .txtTotalCostInclusive.Text = rsCONTROL!TotalCostInclusive
            .txtTotalVATAmount.Text = rsCONTROL!VATAmount & ""
            
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadOrder()
On Error GoTo err
        With Screen.ActiveForm

            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASMPurchaseOrder WHERE OrderNo = '" & .txtOrderNo.Text & "' ;"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
            If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub

            .txtOrderDate.Text = rsCONTROL!OrderDate
            .txtOrderDescription.Text = rsCONTROL!OrderDescription
            
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadPuchaseOrderItems()
On Error GoTo err
        With Screen.ActiveForm

            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASMPurchaseOrderItems WHERE OrderNo = '" & .txtOrderNo.Text & "' and ProductCode = '" & .txtProductCode.Text & "' ;"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
            If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub

            .txtUnitCost.Text = rsCONTROL!UnitCost
            .txtUnitQuantity.Text = rsCONTROL!Balance
            .txtTotalUnitCost.Text = rsCONTROL!TotalCost
            .txtVATAmount.Text = rsCONTROL!VATAmount
            .txtVATRate.Text = rsCONTROL!VATRate
'            .UpDownQuantity = CDbl(.txtUnitQuantity.Text)
            .txtTotalUnitCostInclusive.Text = rsCONTROL!TotalCostInclusive
            
            '/ Used to set the Defaulf value for Quick Data Entry
            
            .txtReceivedUnitCost.Text = rsCONTROL!UnitCost
            .txtReceivedUnitQuantity.Text = rsCONTROL!Balance
            .txtReceivedTotalUnitCost.Text = rsCONTROL!TotalCost
            .txtReceivedVATAmount.Text = rsCONTROL!VATAmount
            .txtReceivedTotalUnitCostInclusive.Text = rsCONTROL!TotalCostInclusive
            .txtReceivedVATRate.Text = rsCONTROL!VATRate

        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadPuchaseOrderRECEIVED()
On Error GoTo err
        With frmODASMProductSetup

            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASMPurchaseOrderItems WHERE OrderNo = '" & .txtOrderNo.Text & "' and ProductCode = '" & .txtProductCode.Text & "' ;"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
            If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub

            '/ Used to set the Defaulf value for Quick Data Entry
            
            .txtReceivedUnitCost.Text = rsCONTROL!UnitCost
            .txtReceivedUnitQuantity.Text = rsCONTROL!Balance
            .txtReceivedTotalUnitCost.Text = rsCONTROL!TotalCost
            .txtReceivedVATAmount.Text = rsCONTROL!VATAmount
            .txtReceivedTotalUnitCostInclusive.Text = rsCONTROL!TotalCostInclusive
            .txtReceivedVATRate.Text = rsCONTROL!VATRate
            .txtUnitType.Text = rsCONTROL!UnitCode & ""
        End With
Exit Sub

err:
    ErrorMessage
End Sub



Public Sub ClearPurchaseOrderItems()
On Error GoTo err
    With Screen.ActiveForm
        .txtProductDescription.Text = Empty
        .txtTotalUnitCost.Text = 0
        .txtTotalUnitCostInclusive.Text = 0
        .txtUnitCost.Text = 0
        .txtUnitQuantity.Text = 0
        .txtVATAmount.Text = 0
    End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub clearGRNItems()
On Error GoTo err
    With frmODASMProductSetup
        .txtReceivedTotalUnitCost.Text = 0
        .txtReceivedTotalUnitCostInclusive.Text = 0
        .txtReceivedUnitCost.Text = 0
        .txtReceivedUnitQuantity.Text = 0
        .txtReceivedVATRate.Text = 0
        .txtReceivedVATAmount.Text = 0
        .txtProductCode.Text = Empty
        .txtProductDescription.Text = Empty
        .txtQuantity.Text = 1
        
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub generateGRNNo()
On Error GoTo err
        Set rsCONTROL = New Recordset
        
        strSQL = "Select * from ODASPLAstNumbers Where AutoGRNNo = 'Y';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF Or .BOF = True Then Exit Sub
                
                frmODASMReceiveOrder.txtGRNNo.Text = !GRNNo
                
                Select Case Len(frmODASMReceiveOrder.txtGRNNo.Text)
                        Case 1: frmODASMReceiveOrder.txtGRNNo.Text = Trim(!GRNPrefix) + "0000" + Trim(frmODASMReceiveOrder.txtGRNNo.Text)
                        Case 2: frmODASMReceiveOrder.txtGRNNo.Text = Trim(!GRNPrefix) + "000" + Trim(frmODASMReceiveOrder.txtGRNNo.Text)
                        Case 3: frmODASMReceiveOrder.txtGRNNo.Text = Trim(!GRNPrefix) + "00" + Trim(frmODASMReceiveOrder.txtGRNNo.Text)
                        Case 4: frmODASMReceiveOrder.txtGRNNo.Text = Trim(!GRNPrefix) + "0" + Trim(frmODASMReceiveOrder.txtGRNNo.Text)
                        Case 5: frmODASMReceiveOrder.txtGRNNo.Text = Trim(!GRNPrefix) + Trim(frmODASMReceiveOrder.txtGRNNo.Text)
                End Select
                                
                !GRNNo = Str(Val(!GRNNo) + 1)
                .Update
                .Requery
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub SaveGRNItems()
On Error GoTo err
        With frmODASMReceiveOrder

                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMGRNItems WHERE GRNNo = '" & .txtGRNNo.Text & "' and ProductCode = '" & .txtProductCode.Text & "' ;"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                    
                            If rsSAVE.EOF And rsSAVE.BOF Then
                                    rsSAVE.AddNew
                                    rsSAVE!GRNNo = .txtGRNNo.Text
                                    rsSAVE!OrderNo = .txtOrderNo.Text
                                    rsSAVE!ProductCode = .txtProductCode.Text
                                    rsSAVE!Prepared = "Y"
                                    rsSAVE!PreparedBy = CurrentUserName
                                    rsSAVE!DatePrepared = Date
                                    rsSAVE!Received = "Y"
                                    rsSAVE!QuantityReceived = 0
                                    rsSAVE!Balance = CDbl(.txtUnitQuantity.Text)
                                    rsSAVE!InventoryUpdated = "N"
                                    
                             End If
                            
                            rsSAVE!UnitQuantity = CDbl(.txtReceivedUnitQuantity.Text)
                            'rsSAVE!CurrencyCode = .cboCurrencyCode.Text
                            rsSAVE!UnitCost = CCur(.txtUnitCost.Text)
                            rsSAVE!Status = "ORDER-RECEIVED"
                            rsSAVE!VATRate = CDbl(.txtReceivedVATRate)
                            rsSAVE!TotalCost = CCur(.txtReceivedTotalUnitCost.Text)
                            rsSAVE!VATAmount = CCur(.txtReceivedVATAmount.Text)
                            'rsSAVE!ExchangeRate = CCur(.txtExchangeRate.Text)
                            rsSAVE!TotalCostInclusive = CCur(.txtReceivedTotalUnitCostInclusive.Text)
                            rsSAVE.Update
                            rsSAVE.Requery
                        

        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub SaveRECORD()
On Error GoTo err
        With frmODASMReceiveOrder

            Set rsSAVE = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASMPurchaseOrderItems WHERE OrderNo = '" & .txtOrderNo.Text & "' and ProductCode = '" & .txtProductCode.Text & "' ;"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
            If rsSAVE.EOF And rsSAVE.BOF Then
                    rsSAVE.AddNew
                    rsSAVE!OrderNo = .txtOrderNo.Text
                    rsSAVE!ProductCode = .txtProductCode.Text
                    rsSAVE!Prepared = "Y"
                    rsSAVE!PreparedBy = CurrentUserName
                    rsSAVE!DatePrepared = Date
                    rsSAVE!QuantityReceived = 0
                    rsSAVE!Balance = CDbl(.txtUnitQuantity.Text)
             End If
            
            rsSAVE!Received = "Y"
            rsSAVE!receivedBy = CurrentUserName
            rsSAVE!DateReceived = .txtDeliveryDate.Text
            rsSAVE!QuantityReceived = CDbl(.txtReceivedUnitQuantity.Text)
            rsSAVE!ReceivedUnitCost = CCur(.txtReceivedTotalUnitCost.Text)
            rsSAVE!Balance = rsSAVE!Balance - rsSAVE!QuantityReceived
            
            If rsSAVE!Balance <= 0 Then
                    rsSAVE!Status = "FULL DELIVERY"
                Else:
                    rsSAVE!Status = "PART DELIVERY"

            End If
            
            bSaveRECORD = False
            rsSAVE.Update
            rsSAVE.Requery
                    

    End With

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub updateRECORD()
        
        
        If bSaveRECORD = True Then
                If frmODASMReceiveOrder.txtGRNNo.Text = Empty Then
                        generateGRNNo
                End If
                
                saveGRN
                SaveGRNItems
                SaveRECORD
                updatePURCHASEORDER
                
                If bSaveRECORD = False Then
                        disableALLRECORD
                        'calculateTOTALS
                End If
        End If
        
        showORDERITEMSWITHBALANCE
        showORDERITEMSRECEIVED

End Sub
Public Sub calculateTOTALS()
On Error GoTo err
        With frmODASMReceiveOrder
                
                '/ Calculate the Total Cost of All the Items in the Purchase Order
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(TotalCost) as TOTAL FROM ODASMPurchaseOrderItems WHERE OrderNo ='" & .txtOrderNo.Text & "'"
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Total) = True Then
                            .txtTotalCost.Text = 0
                Else: .txtTotalCost.Text = FormatCurrency(rsCONTROL!Total)
                End If
                
                strSQL = Empty
                
                '/ * CALCULATE total cost of Items Inclusive
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(TotalCostInclusive) as TOTAL FROM ODASMPurchaseOrderItems WHERE OrderNo = '" & .txtOrderNo.Text & "'"
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Total) = True Then
                            .txtTotalCostInclusive.Text = 0
                Else: .txtTotalCostInclusive.Text = FormatCurrency(rsCONTROL!Total)
                End If
                
                strSQL = Empty

                '/ * CALCULATE TOTAL VAT AMOUNT OF ALL THE ITEMS
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(VATAmount) as TOTAL FROM ODASMPurchaseOrderItems WHERE OrderNo = '" & .txtOrderNo.Text & "' "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Total) = True Then
                            .txtTotalVATAmount.Text = 0
                Else: .txtTotalVATAmount.Text = FormatCurrency(rsCONTROL!Total)
                End If

                
                strSQL = Empty

               '/ Calculate Qty of Items Ordered
                Dim numQUANTITY As Double
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(UnitQuantity) as TOTAL FROM ODASMPurchaseOrderItems WHERE OrderNo = '" & .txtOrderNo.Text & "'"
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Total) = True Then
                            numQUANTITY = 0
                Else: numQUANTITY = FormatNumber(rsCONTROL!Total)
                End If
                
                strSQL = Empty

                '/ * This proedure is used to update the Purchase Order with the Required Details
                
                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMPurchaseOrder WHERE OrderNo = '" & .txtOrderNo.Text & "'; "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsSAVE.EOF And rsSAVE.BOF Then Exit Sub
                
                rsSAVE!TotalCost = CCur(frmODASMReceiveOrder.txtTotalCost.Text)
                rsSAVE!TotalCostInclusive = CCur(frmODASMReceiveOrder.txtTotalCostInclusive.Text)
                rsSAVE!VATAmount = CCur(frmODASMReceiveOrder.txtTotalVATAmount.Text)
                rsSAVE!quantity = CCur(numQUANTITY)
                rsSAVE.Update
                rsSAVE.Requery
                
                
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub ValidateRECORD()
On Error GoTo err
    bSaveRECORD = False
    With frmODASMReceiveOrder
        
        If .txtGRNDate.Text = Empty Then
                MsgBox "The Goods Received note Date Recieved "
                .txtGRNDate.SetFocus
        
        ElseIf CDbl(.txtReceivedUnitQuantity.Text) <= 0 Then
                MsgBox "The Goods received MUST be Greater Than Zero"
                .txtReceivedUnitQuantity.SetFocus
        
        ElseIf CDbl(.txtReceivedTotalUnitCost.Text) <= 0 Then
                MsgBox "The Total cost of the Goods CANNOT be Less Than Zero"
                .txtReceivedTotalUnitCost.SetFocus
            
        ElseIf CDbl(.txtReceivedTotalUnitCostInclusive) <= 0 Then
                MsgBox "The Total cost (Inclusive) of the Goods CANNOT be Less Than Zero"
                .txtReceivedTotalUnitCostInclusive.SetFocus
        
        ElseIf CDbl(.txtReceivedTotalUnitCostInclusive) < CDbl(.txtReceivedTotalUnitCost.Text) Then
                MsgBox "The Total Amount Inclusive cannot be Less than the Total Amount"
                .txtReceivedTotalUnitCostInclusive.SetFocus
        
        ElseIf .txtDeliveryNoteNo.Text = Empty Then
                MsgBox "The Delievry Note Number cannot be Left Blank"
                .txtDeliveryNoteNo.SetFocus
        
        ElseIf .txtDeliveryDate.Text = Empty Then
                MsgBox "The Delivery Date Cannot be Left Blank"
                .txtDeliveryDate.SetFocus
        Else
                bSaveRECORD = True
        End If

    End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub saveGRN()
On Error GoTo err

        With frmODASMReceiveOrder

                If rsSAVE.EOF And rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!GRNNo = .txtGRNNo.Text
                        rsSAVE!GrnDate = .txtGRNDate.Text
                        rsSAVE!OrderNo = .txtOrderNo.Text
                        rsSAVE!Prepared = "Y"
                        rsSAVE!PreparedBy = CurrentUserName
                        rsSAVE!DatePrepared = Date
                        rsSAVE!TotalCost = 0
                        rsSAVE!CurrentPeriod = CurrentPeriod
                        rsSAVE!InventoryUpdated = "N"
                End If
                
                If rsSAVE!DeliveryNoteNo = Empty Then
                        rsSAVE!Delivered = "N"
                Else: rsSAVE!Delivered = "Y"
                        rsSAVE!DeliveryDate = .txtDeliveryDate.Text
                        rsSAVE!DeliveryNoteNo = .txtDeliveryNoteNo.Text
                End If
                
                rsSAVE!supplierCode = .txtSupplierCode.Text
                rsSAVE!Remarks = .txtRemarks.Text
                rsSAVE!OrderDate = .txtOrderDate.Text
                rsSAVE!DateRequired = .txtDeadlineDate.Text
                rsSAVE.Update
                rsSAVE.Requery
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub CalculateMARKUP()
On Error GoTo err

    With frmODASMProductSetup
            
            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASPProduct WHERE ProductCode = '" & .txtProductCode.Text & "';  "
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsCONTROL.BOF Or rsCONTROL.EOF Then Exit Sub
            
            If .txtQuantity.Text = Empty Then Exit Sub
            
            .txtMarkupType.Text = Trim(rsCONTROL!Markuptype) & ""
            .txtMarkUp.Text = CDbl(rsCONTROL!Markup)
            .txtTotalQuantity.Text = CDbl(.txtReceivedUnitQuantity.Text) * CDbl(.txtQuantity)

            .txtProductCost.Text = CDbl(.txtReceivedTotalUnitCost) / CDbl(.txtTotalQuantity.Text)
            
            '/ Procedure to calculate Markup
            If .txtMarkupType.Text = "P" Then
                    .txtMarkupAmount = CDbl(.txtProductCost.Text) * CDbl(.txtMarkUp) / 100
            Else: .txtMarkupAmount = CDbl(.txtMarkUp)
            End If
            
            .txtProductPrice.Text = CDbl(.txtProductCost.Text) + CDbl(.txtMarkupAmount.Text)

    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub updatePURCHASEORDER()
On Error GoTo err

        With frmODASMReceiveOrder

                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMPurchaseOrder WHERE OrderNo = '" & .txtOrderNo.Text & "';  "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        
                If rsSAVE.EOF And rsSAVE.BOF Then Exit Sub
                rsSAVE!Received = "P"
                rsSAVE!DateReceived = Date
                rsSAVE!receivedBy = CurrentUserName
                rsSAVE.Update
                rsSAVE.Requery
    End With

Exit Sub

err:
    ErrorMessage
End Sub

