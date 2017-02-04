VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALISPaymentRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub clearRECORD()
On Error GoTo err

    With frmODASMVoucher
        .txtStatus.Text = ""
        .cboDocumentNo.Text = ""
        .txtReference.Text = ""
        .txtPaymentDescription.Text = ""
        .txtPaymentDescription.Text = ""
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadPAYEE()
On Error GoTo err
            frmALISMPayee.txtPayeeDetails = strPAYEE
            frmALISMPayee.txtPayeeAddress = strPOSTALADDRESS
            frmALISMPayee.txtPostalCode = strPOSTALCODE
            frmALISMPayee.cboTownCode = strTOWN

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadPaymentGRID()
On Error GoTo err

    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub
    If bloadClaimProceeds Then
            strSQL = "SELECT ALISMClaim.ClaimType, ALISPClaimType.ClaimTypeDescription, ALISMClaim.Amount FROM ALISMClaim, ALISPClaimType WHERE ALISMClaim.ClaimType = ALISPClaimType.ClaimType AND ALISMClaim.claimNo =  '" & frmODASMVoucher.cboDocumentNo & "' and ALISMClaim.type = 'A' ;"
    ElseIf bloadClaimDeductions Then
            strSQL = "SELECT ALISMClaim.ClaimType, ALISPClaimType.ClaimTypeDescription, ALISMClaim.Amount FROM ALISMClaim, ALISPClaimType WHERE ALISMClaim.ClaimType = ALISPClaimType.ClaimType AND ALISMClaim.claimNo =  '" & frmODASMVoucher.cboDocumentNo & "' and ALISMClaim.type = 'D' ;"
    End If
    
    GetProceeds
    
    
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub GetProceeds()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Type", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Particulars", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ClaimType))
                        
                        If Not IsNull(rsLIST!ClaimTypeDescription) Then
                            MyList.SubItems(1) = CStr(rsLIST!ClaimTypeDescription)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                            MyList.SubItems(2) = CStr(rsLIST!Amount)
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetCostCenter()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Payment Type", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Payment Description", .ListView1.Width / 2
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CostCenter))
                        
                        If Not IsNull(rsLIST!Description) Then
                            MyList.SubItems(1) = CStr(rsLIST!Description)
                        End If
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub
Public Sub GetInvoicesNotPaid()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Payment Code", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Payment Description", .ListView1.Width / 2
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PaymentCode))
                        
                        If Not IsNull(rsLIST!PaymentCodeDescription) Then
                            MyList.SubItems(1) = CStr(rsLIST!PaymentCodeDescription)
                        End If
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

