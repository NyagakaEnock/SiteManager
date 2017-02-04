VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub LoadDEFAULT()
On Error GoTo err
        
        With frmODASMOrder
            .txtOrderDate.Text = Date
            .txtDeadlineDate.Text = Date
            .txtUnitQuantity.Text = 1
            .txtUnitCost.Text = 0
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub obtainBASECURRENCY()
On Error GoTo err
            
            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASPCurrency WHERE BaseCurrency = 'Y'; "
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
            Screen.ActiveForm.txtExchangeRate.Text = 1
            Screen.ActiveForm.cboCurrencyCode.Text = rsCONTROL!CurrencyCode
            Screen.ActiveForm.txtCurrencySymbol.Text = rsCONTROL!CurrencySymbol

rsCONTROL.Close
strSQL = Empty
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub SaveRECORD()
On Error GoTo err
    With frmODASMOrder

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
                        rsSAVE!Received = "N"
                        rsSAVE!QuantityReceived = 0
                        rsSAVE!Balance = CDbl(.txtUnitQuantity.Text)
                 End If
                
                rsSAVE!UnitQuantity = CDbl(.txtUnitQuantity.Text)
                rsSAVE!Balance = rsSAVE!Balance
                rsSAVE!UnitCode = .cboUnitCode.Text
                rsSAVE!CurrencyCode = .cboCurrencyCode.Text
                rsSAVE!UnitCost = CCur(.txtUnitCost.Text)
                rsSAVE!Status = "REQ-PLACED"
                rsSAVE!VATRate = CDbl(.txtVATRate)
                rsSAVE!TotalCost = CCur(.txtTotalUnitCost.Text)
                rsSAVE!VATAmount = CCur(.txtVATAmount.Text)
                rsSAVE!ExchangeRate = CCur(.txtExchangeRate.Text)
                rsSAVE!TotalCostInclusive = CCur(.txtTotalUnitCostInclusive.Text)
                bSaveRECORD = False
                rsSAVE.Update
                rsSAVE.Requery
            

        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub saveORDER()
On Error GoTo err

        With frmODASMOrder

                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMPurchaseOrder WHERE OrderNo = '" & .txtOrderNo.Text & "';  "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        
                If rsSAVE.EOF And rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!OrderNo = .txtOrderNo.Text
                        rsSAVE!Prepared = "Y"
                        rsSAVE!PreparedBy = CurrentUserName
                        rsSAVE!DatePrepared = Date
                        rsSAVE!TotalCost = 0
                        rsSAVE!CurrentPeriod = CurrentPeriod
                 End If
                
                rsSAVE!supplierCode = .txtSupplierCode.Text
                rsSAVE!Remarks = .txtRemarks.Text
                rsSAVE!OrderDate = .txtOrderDate.Text
                rsSAVE!OrderDescription = Trim(.txtOrderDescription.Text)
                rsSAVE!DateRequired = .txtDeadlineDate.Text
                rsSAVE.Update
                rsSAVE.Requery
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub saveISSUED()
On Error GoTo err

        With frmODASMIssueOrder

                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMPurchaseOrder WHERE OrderNo = '" & .txtOrderNo.Text & "';  "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        
                If rsSAVE.EOF And rsSAVE.BOF Then Exit Sub
                 
                If .optIssue.Value = True Then
                        rsSAVE!Issued = "Y"
                        rsSAVE!DateIssued = .txtDateIssued
                        rsSAVE!IssuedBy = .txtIssueBy
                
                ElseIf .optCancel.Value = True Then
                        rsSAVE!Cancelled = "Y"
                        rsSAVE!DateCancelled = .txtDateCancelled.Text
                        rsSAVE!CancelledBy = .txtCancelledBy.Text
                End If
                
                bSaveRECORD = False
                
                rsSAVE.Update
                rsSAVE.Requery
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadISSUED()
On Error GoTo err

        With frmODASMIssueOrder

                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMPurchaseOrder WHERE OrderNo = '" & .txtOrderNo.Text & "';  "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                .txtOrderDate.Text = rsCONTROL!OrderDate
                .txtDeadlineDate.Text = rsCONTROL!DateRequired
                .txtRemarks.Text = rsCONTROL!Remarks & ""
                .txtSupplierCode.Text = rsCONTROL!supplierCode
                .txtTotalCost.Text = rsCONTROL!TotalCost
                .txtTotalCostInclusive.Text = rsCONTROL!TotalCostInclusive
                .txtTotalVATAmount.Text = rsCONTROL!VATAmount & ""
    End With

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub CheckDetails()
On Error GoTo err
    If NewRecord = True Or beditRECORD = True Then
        With frmODASMIssueOrder
                         
             
            .txtDateIssued.Text = Empty
            .txtIssueBy.Text = Empty
            .txtCancelledBy.Text = Empty
            .txtDateCancelled.Text = Empty

            If .optIssue.Value = True Then
                .txtDateIssued.Text = Date
                .txtIssueBy.Text = CurrentUserName
            ElseIf .optCancel.Value = True Then
                .txtCancelledBy.Text = CurrentUserName
                .txtDateCancelled.Text = Date
            End If
        End With
        
    End If

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub validIssue()
On Error GoTo err
    With frmODASMIssueOrder
        
        bSaveRECORD = True
        
        If .txtOrderNo.Text = Empty Then
                MsgBox "Cannot Process this Record, Kindly Select the Appropriate Order"
                .txtOrderNo.SetFocus
        
        ElseIf .txtCancelledBy.Text = Empty And .txtIssueBy.Text = Empty Then
                MsgBox "Must be Either Issued by or cancelled By"
                .txtIssueBy.SetFocus
        
        ElseIf .txtDateCancelled.Text = Empty And .txtDateIssued.Text = Empty Then
                MsgBox "The Date Issued or Date cancelled are Required"
                .txtDateIssued.SetFocus
        
        ElseIf .optCancel.Value = 0 And .optIssue.Value = 0 Then
                MsgBox "You must select either Issued or Cancelled"
                .optIssue.SetFocus
        Else
                bSaveRECORD = True
        End If
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub ValidDateRecord()
On Error GoTo err
        bSaveRECORD = False
        With frmODASMOrder
            
            If .txtOrderNo.Text <= "" And NewRecord <> True Then
                    MsgBox "The Order Number Cannot be Blank"
                    .txtOrderNo.SetFocus
            
            ElseIf .cboCurrencyCode.Text = Empty Then
                    MsgBox "The Currency Code cannot be Left Blank"
                    .cboCurrencyCode.SetFocus
            
            ElseIf .txtDeadlineDate.Text = Empty Then
                    MsgBox "The Deadline Date cannot be Left Blank"
                    .txtDeadlineDate.SetFocus
            
            ElseIf CDbl(.txtExchangeRate.Text) = Empty Then
                    MsgBox "The Exchange Rate MUST > 0"
                    .txtExchangeRate.SetFocus

            ElseIf .txtProductCode.Text = Empty Then
                    MsgBox "The product Code Cannot be Left Blank"
                    .txtProductCode.SetFocus
            
            ElseIf .cboUnitCode.Text <= Empty Then
                    MsgBox "The Unit Size Cannot be Left Blank"
                    .cboUnitCode.SetFocus
                    
            ElseIf CDbl(.txtUnitQuantity.Text) = 0 Then
                    MsgBox "The Quantity MUST be Greater Than Zero ........"
                    .txtUnitQuantity.SetFocus
            
            ElseIf CCur(.txtUnitCost.Text) = 0 Then
                    MsgBox "The Cost of the Item is Required ......."
                    .txtUnitCost.SetFocus
            
            ElseIf .txtOrderDescription.Text = Empty Then
                    MsgBox "The Description of the Order is Required ........"
                    .txtOrderDescription.SetFocus
            
            ElseIf .txtOrderDate.Text = Empty Then
                    MsgBox "The Order Date Cannot be Left Blank ........."
                    .txtOrderDate.SetFocus
            
            Else
                    bSaveRECORD = True
            End If
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub generateOrderNO()
On Error GoTo err
        Set rsCONTROL = New Recordset
        
        strSQL = "Select * from ODASPLAstNumbers Where AutoOrderNo = 'Y';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF Or .BOF = True Then Exit Sub
                
                frmODASMOrder.txtOrderNo.Text = !OrderNo
                
                Select Case Len(frmODASMOrder.txtOrderNo.Text)
                        Case 1: frmODASMOrder.txtOrderNo.Text = Trim(!OrderPrefix) + "0000" + Trim(frmODASMOrder.txtOrderNo.Text)
                        Case 2: frmODASMOrder.txtOrderNo.Text = Trim(!OrderPrefix) + "000" + Trim(frmODASMOrder.txtOrderNo.Text)
                        Case 3: frmODASMOrder.txtOrderNo.Text = Trim(!OrderPrefix) + "00" + Trim(frmODASMOrder.txtOrderNo.Text)
                        Case 4: frmODASMOrder.txtOrderNo.Text = Trim(!OrderPrefix) + "0" + Trim(frmODASMOrder.txtOrderNo.Text)
                        Case 5: frmODASMOrder.txtOrderNo.Text = Trim(!OrderPrefix) + Trim(frmODASMOrder.txtOrderNo.Text)
                End Select
                                
                !OrderNo = Str(Val(!OrderNo) + 1)
                .Update
                .Requery
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub ClearPurchaseOrderItems()
On Error GoTo err
        
        With frmODASMOrder
            .txtProductCode.Text = Empty
            .txtProductDescription.Text = Empty
            .txtTotalUnitCost.Text = 0
            .txtTotalUnitCostInclusive.Text = 0
            .txtUnitCost.Text = 0
            .txtUnitQuantity.Text = 0
        End With
        
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub ClearIssued()
On Error GoTo err
        
        With frmODASMIssueOrder
            .txtCancelledBy.Text = Empty
            .txtDateCancelled.Text = Empty
            .optCancel.Value = 0
            .txtIssueBy.Text = CurrentUserName
            .txtDateIssued.Text = Date
            .optIssue.Value = 0
        End With
        
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub updateRECORD()
        
        
        If bSaveRECORD = True Then
                If frmODASMOrder.txtOrderNo.Text = Empty Then
                        generateOrderNO
                End If
                
                saveORDER
                SaveRECORD
                
                If bSaveRECORD = False Then
                        disableALLRECORD
                        calculateTOTALS
                End If
        End If
        
        listALLProducts
        showORDERITEMS
        ShowALLOrdersRequisitioned
End Sub
Public Sub updateISSUED()
        
        
        If bSaveRECORD = True Then
                saveISSUED
                If bSaveRECORD = False Then
                        disableALLRECORD
                End If
        End If
        showORDERITEMS
End Sub

Public Sub calculateTOTALS()
On Error GoTo err
        With frmODASMOrder
                
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
                
                rsSAVE!TotalCost = CCur(frmODASMOrder.txtTotalCost.Text)
                rsSAVE!TotalCostInclusive = CCur(frmODASMOrder.txtTotalCostInclusive.Text)
                rsSAVE!VATAmount = CCur(frmODASMOrder.txtTotalVATAmount.Text)
                rsSAVE!quantity = CCur(numQUANTITY)
                rsSAVE.Update
                rsSAVE.Requery
                
                
        End With
Exit Sub

err:
    ErrorMessage
End Sub

