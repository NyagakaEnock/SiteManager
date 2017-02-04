VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALISBankerOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub getBANKORDER()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Order No", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Policy No", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Order Amount", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Payment Method", .ListView1.Width / 2
                            
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select orderNo, documentNo, OrderDate, OrderAmount, PaymentMethod from ALISMOrders"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!OrderNo))
                        
                        If Not IsNull(rsLIST!DocumentNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!DocumentNo)
                        End If

                        If Not IsNull(rsLIST!OrderDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!OrderDate)
                        End If
                        
                        If Not IsNull(rsLIST!OrderAmount) Then
                            MyList.SubItems(3) = CStr(rsLIST!OrderAmount)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentMethod) Then
                            MyList.SubItems(4) = CStr(rsLIST!PaymentMethod)
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

Public Sub getPASTBANKORDER()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                .ListView2.ColumnHeaders.Add , , "Order No", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Policy No", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Order Date", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Order Amount", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Payment Method", .ListView2.Width / 5
                            
                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select orderNo, documentNo, OrderDate, OrderAmount, PaymentMethod from ALISMOrders where DocumentNo = '" & Screen.ActiveForm.txtAccountNo.Text & "'"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!OrderNo))
                        
                        If Not IsNull(rsLIST!DocumentNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!DocumentNo)
                        End If

                        If Not IsNull(rsLIST!OrderDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!OrderDate)
                        End If
                        
                        If Not IsNull(rsLIST!OrderAmount) Then
                            MyList.SubItems(3) = CStr(rsLIST!OrderAmount)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentMethod) Then
                            MyList.SubItems(4) = CStr(rsLIST!PaymentMethod)
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
Public Sub loadRECORD()
    
    loadPolicy
    loadPAYMENTMETHOD
    loadPaymentModeDESCRIPTION
    loadBANK
    LoadCoyBank
    getBANKORDER
    getPASTBANKORDER
End Sub
Private Sub LoadCoyBank()
On Error GoTo err

        Dim rsBANKLF As ADODB.Recordset, strBANKLF As String
        Set rsBANKLF = New Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBankAccount WHERE BankNo = '" & Screen.ActiveForm.cboCoyBankNO.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.txtCoyBankName.Text = !Details
        End With
        
rsBANKLF.Close
strBANKLF = ""

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub loadPolicy()
On Error GoTo err
        Dim rsPOLICY As ADODB.Recordset
        Set rsPOLICY = New ADODB.Recordset
        
        strSQL = "SELECT * FROM ODASMJobBrief, ODASPAccount WHERE ODASMJobBrief.JobBriefNo= '" & Screen.ActiveForm.txtJobBriefNo.Text & "' and ODASMJobBrief.AccountNo LIKE ODASPAccount.AccountNo ; "
        rsPOLICY.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPOLICY
                If .EOF And .BOF Then Exit Sub
                
                Screen.ActiveForm.txtJobBriefNo.Text = !JobBriefNo
                Screen.ActiveForm.txtNames.Text = Trim(!OtherNames) + " " + Trim(!CompanyName)
                Screen.ActiveForm.cboPaymentMethod.Text = !PaymentMethod
                Screen.ActiveForm.txtDateofFirstPayment.Text = !DateOfCommencement
                Screen.ActiveForm.cboPaymentMode.Text = !PaymentMode
                
                If baddRECORD = True Then
                        Screen.ActiveForm.txtIssuedBy.Text = UserName
                        Screen.ActiveForm.txtOrderDate.Text = Date
                        Screen.ActiveForm.txtOrderAmount.Text = !ExpectedPremium
                        Screen.ActiveForm.txtAccountNo.Text = !AccountNo & ""
                        Screen.ActiveForm.txtBankNo.Text = !BankNo & ""

                End If
       
       End With

rsPOLICY.Close

Exit Sub

err:
        ErrorMessage

End Sub
Public Sub loadORDER()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT * FROM ALISMOrders WHERE OrderNo= '" & Screen.ActiveForm.txtOrderNo.Text & "'; "
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                
                Screen.ActiveForm.txtJobBriefNo.Text = !DocumentNo
                Screen.ActiveForm.cboPaymentMethod.Text = !PaymentMethod
                Screen.ActiveForm.txtAccountNo.Text = !AccountNo & ""
                Screen.ActiveForm.txtIssuedBy.Text = !IssuedBy & ""
                Screen.ActiveForm.cboPaymentMode.Text = !PaymentMode
                Screen.ActiveForm.cboBankName.Text = !BankNo & ""
                Screen.ActiveForm.txtCoyAccountNo.Text = !CoyAccountNo & ""
                Screen.ActiveForm.txtBankNo.Text = !BankNo & ""
                Screen.ActiveForm.txtOrderAmount.Text = !OrderAmount & ""
                Screen.ActiveForm.txtOrderDate.Text = !OrderDate & ""
                Screen.ActiveForm.cboCoyBankNO.Text = !CoyBankNo & ""
       End With

rsCONTROL.Close

Exit Sub

err:
        ErrorMessage

End Sub

Public Sub saveORDER()

        validateRECORD
        If bsaveRECORD = True Then
            saveRecord
            savePOLICY
            
            getBANKORDER
            getPASTBANKORDER
            
            bsaveRECORD = False
            disableALLRECORD
        End If
End Sub
Private Sub validateRECORD()
On Error GoTo err
    
    With Screen.ActiveForm
        If .txtJobBriefNo.Text <= "" Then
                MsgBox " The Policy No cannot be Left Blank", vbOKOnly
                .txtJobBriefNo.SetFocus
        
        ElseIf .txtOrderAmount.Text <= "" Then
                MsgBox "The Amount cannot be Left Blank", vbOKOnly
                .txtOrderAmount.SetFocus
                
        ElseIf .txtOrderDate.Text <= "" Then
                MsgBox "The Order date cannot be Left Blank", vbOKOnly
                .txtOrderDate.SetFocus
        
        ElseIf .cboPaymentMethod.Text <= "" Then
                MsgBox "The Payment Method Cannot be Left Blank", vbOKOnly
                .cboPaymentMethod.SetFocus
        
        ElseIf .cboBankName.Text <= "" Then
                MsgBox "The Policyholders Bank No is Needed for Processing", vbOKOnly
                .cboBankName.SetFocus
                
        ElseIf .txtDateofFirstPayment.Text <= "" Then
                MsgBox "The First Payment Date is Needed", vbOKOnly
                .txtDateofFirstPayment.SetFocus
        
        ElseIf .cboPaymentMode.Text <= "" Then
                MsgBox "The Payment Mode is Required "
                .cboPaymentMode.SetFocus
        
        ElseIf .cboCoyBankNO.Text <= "" Then
                MsgBox "The Companys Bank Account Number is Needed", vbOKOnly
                .cboCoyBankNO.SetFocus
        
        ElseIf .txtCoyAccountNo.Text <= "" Then
                MsgBox "The Companys Bank Account Number is Needed", vbOKOnly
                .txtCoyAccountNo.SetFocus
        
        ElseIf .txtIssuedBy.Text <= "" Then
                MsgBox "The Details of the Person Issuing the Instructions is Required", vbOKOnly
                .txtIssuedBy.SetFocus
        Else
                bsaveRECORD = True
        
        End If
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub saveRecord()
On Error GoTo err

    Set rsSAVE = New ADODB.Recordset
    rsSAVE.Open "SELECT * FROM ALISMOrders where orderNo = '" & Screen.ActiveForm.txtOrderNo.Text & "' ; ", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsSAVE
            If .EOF Or .BOF Then
                    .AddNew
                    !OrderNo = Screen.ActiveForm.txtOrderNo.Text
            End If
            
            !DocumentNo = Screen.ActiveForm.txtJobBriefNo.Text
            !OrderAmount = CDbl(Screen.ActiveForm.txtOrderAmount.Text)
            !OrderDate = Screen.ActiveForm.txtOrderDate.Text
            !PaymentMethod = Screen.ActiveForm.cboPaymentMethod
            !BankNo = Screen.ActiveForm.txtBankNo.Text
            !AccountNo = Screen.ActiveForm.txtAccountNo.Text
            
            figures = Screen.ActiveForm.txtOrderAmount
            Set translater = New cMoneyConverter
            translater.CallConverter
            !inwords = inwords
            Set translater = Nothing
            
            !DateofFirstPayment = Screen.ActiveForm.txtDateofFirstPayment.Text
            !PaymentMode = Screen.ActiveForm.cboPaymentMode.Text
            !CoyBankNo = Screen.ActiveForm.cboCoyBankNO.Text
            !CoyAccountNo = Screen.ActiveForm.txtCoyAccountNo.Text
            !IssuedBy = Screen.ActiveForm.txtIssuedBy.Text
            !DocumentType = "POL"
            !Preparedby = CurrentUserName
            !dateprepared = Date
            .Update
            .Requery
            
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub printRECORD()
On Error GoTo err
    If Screen.ActiveForm.txtOrderNo.Text <= "" Then
        MsgBox "Cannot Print The Details, Kindly Load the Bankers Other", vbOKOnly
        Screen.ActiveForm.txtOrderNo.SetFocus
    ElseIf Screen.ActiveForm.txtJobBriefNo.Text <= "" Then
        MsgBox "The Policy No Does not Exist", vbOKOnly
    Else
        Load frmALISRBankersOrder
        frmALISRBankersOrder.Show 1
    End If
    
Exit Sub
err:
    ErrorMessage
End Sub


Public Sub addRECORD()
        enableALLRECORD
End Sub

Public Sub Cancelrecord()
        clearALLRECORD
        disableALLRECORD
        enableButtons
End Sub
Private Sub savePOLICY()
On Error GoTo err

        Dim rsPOLICY As ADODB.Recordset
        Set rsPOLICY = New ADODB.Recordset
      
        rsPOLICY.Open "SELECT * FROM ODASMJobBrief WHERE ODASMJobBrief.JobBriefNo= '" & Screen.ActiveForm.txtJobBriefNo.Text & "' ; ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPOLICY
                If .EOF And .BOF Then Exit Sub
                !PaymentMethod = Screen.ActiveForm.cboPaymentMethod.Text
                !AccountNo = Screen.ActiveForm.txtAccountNo.Text
                !BankNo = Screen.ActiveForm.txtBankNo.Text
                .Update
                .Requery
        End With
        
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadBANK()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        rsCONTROL.Open "SELECT * FROM ALISPBank WHERE BankNo = '" & Screen.ActiveForm.txtBankNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                 Screen.ActiveForm.cboBankName.Text = !CompanyName
        End With
        
rsCONTROL.Close
strBANKLF = ""

Exit Sub

err:
        ErrorMessage

End Sub


