Attribute VB_Name = "ModSelection"
Option Explicit

Public Sub SelectSettlementModeGotFocus()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPPaymentMode;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboSettlementMode.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboSettlementMode.AddItem !Description
                    .MoveNext
            Loop
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectSettlementModeLostFocus()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPPaymentMode WHERE Description= '" & Screen.ActiveForm.cbo.cboSettlementMode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboSettlementMode.Text = !PaymentMode
        End With
        
Exit Sub

rsCONTROL.Close

err:
        ErrorMessage

End Sub

Public Sub LoadSettlementDescription()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" & Screen.ActiveForm.cboSettlementMode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboSettlementMode.Text = !SettlementMode
        End With
        
Exit Sub

rsCONTROL.Close

err:
        ErrorMessage

End Sub

Public Sub SelectDurationCodeGotFocus()
On Error GoTo err

        Dim rsDurationCode As ADODB.Recordset, strDurationCode As String
        Set rsDurationCode = New Recordset
      
        strDurationCode = "SELECT * FROM ODASPDuration;"
        rsDurationCode.Open strDurationCode, cnCOMMON, adOpenKeyset, adLockOptimistic

        Screen.ActiveForm.cboDurationMode.Clear

            With rsDurationCode
                    Do Until .EOF
                            Screen.ActiveForm.cboDurationMode.AddItem !DurationDescription
                            .MoveNext
                    Loop
            End With

rsDurationCode.Close
Exit Sub

err:
   ErrorMessage
End Sub
Public Sub SelectDurationCodeKeyPress(KeyAscii As Integer)
         KeyAscii = 0
End Sub

Public Sub SelectDurationCodeLostFocus()
On Error GoTo err

        Dim rsDurationCode As ADODB.Recordset
        Set rsDurationCode = New Recordset
        
        rsDurationCode.Open "SELECT * FROM ODASPDuration WHERE DurationDescription = '" & Screen.ActiveForm.cboDurationMode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsDurationCode
                If .EOF And .BOF Then Exit Sub
                     Screen.ActiveForm.cboDurationMode.Text = !DurationMode
        End With
        
Exit Sub

rsDurationCode.Close

err:
   ErrorMessage

End Sub

Public Sub selectPaymentMethodGotFocus()
On Error GoTo err

        Set rsCONTROL = New Recordset
      
        strSQL = "SELECT * FROM ODASPPaymentMethod;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        Screen.ActiveForm.cboPaymentMethod.Clear

            With rsCONTROL
                    Do Until .EOF
                            Screen.ActiveForm.cboPaymentMethod.AddItem !PaymentMethodDescription
                            .MoveNext
                    Loop
            End With

rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub selectGuarantorGotFocus()
On Error GoTo err

        Set rsCONTROL = New Recordset
      
        strSQL = "SELECT * FROM ODASPGuarantor;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        Screen.ActiveForm.cboGuarantorType.Clear

            With rsCONTROL
                    Do Until .EOF
                            Screen.ActiveForm.cboGuarantorType.AddItem !Guarantor
                            .MoveNext
                    Loop
            End With

rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub selectPaymentMethodKeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Public Sub selectPaymentMethod1LostFocus()
On Error GoTo err

        Dim rsPAYMETHOD As ADODB.Recordset, strPAYMETHOD As String
        Set rsPAYMETHOD = New Recordset
        
        rsPAYMETHOD.Open "SELECT * FROM ODASPPaymentMethod WHERE PaymentMethodDescription= '" & Screen.ActiveForm.cboPaymentMethod.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPAYMETHOD
                If .EOF And .BOF Then Exit Sub
                        Screen.ActiveForm.cboPaymentMethod.Text = !PaymentMethod
        End With
        
rsPAYMETHOD.Close
strPAYMETHOD = ""

Exit Sub

err:
    ErrorMessage

End Sub

Public Sub selectGuarantorLostFocus()
On Error GoTo err

        Dim rsPAYMETHOD As ADODB.Recordset, strPAYMETHOD As String
        Set rsPAYMETHOD = New Recordset
        
        rsPAYMETHOD.Open "SELECT * FROM ODASPGuarantor WHERE Guarantor = '" & Screen.ActiveForm.cboGuarantorType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPAYMETHOD
                If .EOF And .BOF Then Exit Sub
                        Screen.ActiveForm.cboGuarantorType.Text = !GuarantorType
        End With
        
rsPAYMETHOD.Close
strPAYMETHOD = ""

Exit Sub

err:
    ErrorMessage

End Sub
Public Sub loadCostCenter()
On Error GoTo err
    With frmODASMCheck
    
            Dim rsCOSTCENTER As ADODB.Recordset, strCOSTCENTER As String
            Set rsCOSTCENTER = New Recordset
            
            rsCOSTCENTER.Open "SELECT * FROM ODASPCOSTCENTRe WHERE COSTCENTRE= '" & .txtCostCenter.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    
            If rsCOSTCENTER.EOF And rsCOSTCENTER.BOF Then Exit Sub
            .txtCostCenterDescription.Text = rsCOSTCENTER!COSTCENTREDescription
    
    End With
        
rsCOSTCENTER.Close
strCOSTCENTER = ""

Exit Sub

err:
    ErrorMessage

End Sub

Public Sub loadPAYMENTMETHOD()
On Error GoTo err

        Dim rsPAYMETHOD As ADODB.Recordset, strPAYMETHOD As String
        Set rsPAYMETHOD = New Recordset
        
        rsPAYMETHOD.Open "SELECT * FROM ODASPPaymentMethod WHERE PaymentMethod= '" & Screen.ActiveForm.cboPaymentMethod.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPAYMETHOD
                If .EOF And .BOF Then Exit Sub
                        Screen.ActiveForm.txtPaymentMethodDescription.Text = !PaymentMethodDescription
        End With
        
rsPAYMETHOD.Close
strPAYMETHOD = ""

Exit Sub

err:
    ErrorMessage

End Sub

Public Sub loadGuarantor()
On Error GoTo err

        Dim rsPAYMETHOD As ADODB.Recordset, strPAYMETHOD As String
        Set rsPAYMETHOD = New Recordset
        
        rsPAYMETHOD.Open "SELECT * FROM ODASPGuarantor WHERE GuarantorType = '" & Screen.ActiveForm.cboGuarantorType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPAYMETHOD
                If .EOF And .BOF Then Exit Sub
                        Screen.ActiveForm.txtGuarantor.Text = !Guarantor
        End With
        
rsPAYMETHOD.Close
strPAYMETHOD = ""

Exit Sub

err:
    ErrorMessage

End Sub


Public Sub GetPaymentCode()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Payment Code", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Cost Center", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select PaymentCode,CostCenter, PaymentCodeDescription  from ODASPPaymentCode ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PaymentCode))
                        
                        If Not IsNull(rsLIST!CostCenter) Then
                                MyList.SubItems(1) = CStr(rsLIST!CostCenter)
                        End If

                        If Not IsNull(rsLIST!PaymentCodeDescription) Then
                                MyList.SubItems(2) = CStr(rsLIST!PaymentCodeDescription)
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

Public Sub selectPaymentModeGotFocus()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPPaymentMode WHERE Active='Y';"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboPaymentMode.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboPaymentMode.AddItem !PaymentModeDescription
                    .MoveNext
            Loop
    
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectDurationModeGotFocus()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPPaymentMode;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboPaymentMode.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboPaymentMode.AddItem !PaymentModeDescription
                    .MoveNext
            Loop
    
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectPaymentModeLostFocus()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentModeDescription = '" & Screen.ActiveForm.cboPaymentMode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboPaymentMode.Text = !PaymentMode
                'Screen.ActiveForm.txtPaymentModeDescription.Text = !PaymentModeDescription
        End With
  
rsCONTROL.Close

Exit Sub

err:
        ErrorMessage

End Sub

Public Sub selectPayModeLostFocus()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPPaymentMode WHERE Description = '" & Screen.ActiveForm.cboPaymentMode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboPaymentMode.Text = !PaymentMode
        End With
  
rsCONTROL.Close

Exit Sub

err:
        ErrorMessage

End Sub

Public Sub SelectIDTypeLostFocus()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        rsCONTROL.Open "SELECT * FROM ALISPIDType WHERE IDTypeDescription = '" & Screen.ActiveForm.cboIdType.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                            Screen.ActiveForm.cboIdType.Text = !IDType
        End With
        
        Exit Sub

err:
   ErrorMessage

End Sub

Public Sub SelectIDTypeGotFocus()
On Error GoTo err

        Set rsCONTROL = New Recordset
      
        strSQL = "SELECT * FROM ALISPIDType;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        Screen.ActiveForm.cboIdType.Clear

            With rsCONTROL
                    Do Until .EOF
                            Screen.ActiveForm.cboIdType.AddItem !IDTypeDescription
                            .MoveNext
                    Loop
            End With
        
        Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadPaymentModeDESCRIPTION()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" & Screen.ActiveForm.cboPaymentMode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.txtPaymentModeDescription.Text = !Description
        End With
  
rsCONTROL.Close

Exit Sub

err:
        ErrorMessage

End Sub
