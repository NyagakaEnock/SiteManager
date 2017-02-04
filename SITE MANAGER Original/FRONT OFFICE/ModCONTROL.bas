Attribute VB_Name = "ModCONTROL"
Public Sub disableButtons()
On Error GoTo err
            With Screen.ActiveForm
                    .cmdUpdate.Enabled = True
                    .cmdadd.Enabled = False
                    .cmdSearch.Enabled = False
                    .cmdEdit.Enabled = False
                    .cmdDelete.Enabled = False
                    .cmdCancel.Enabled = True
           End With
Exit Sub

err:
        ErrorMessage
End Sub


Public Sub enableButtons()
On Error GoTo err
            With Screen.ActiveForm
                    .cmdUpdate.Enabled = False
                    .cmdadd.Enabled = True
                    .cmdSearch.Enabled = True
                    .cmdEdit.Enabled = True
                    .cmdDelete.Enabled = True
                    .cmdCancel.Enabled = True
            End With
            Exit Sub
err:
ErrorMessage

End Sub

Public Sub ProductCodeGotFocus()
On Error GoTo err:

        Dim rsPRODUCT As ADODB.Recordset, strPRODUCT As String
        Set rsPRODUCT = New Recordset
      
        If banticipatedENDOWMENT = True Then
                strPRODUCT = "SELECT * FROM ALISPProduct, ALISPPaidup Where ALISPProduct.productCode = ALISPPaidup.productCode and ALISPPaidup.AnticipatedEndowment = 'Y';"
        Else: strPRODUCT = "SELECT * FROM ALISPProduct;"
        End If
        rsPRODUCT.Open strPRODUCT, cnALIS, adOpenKeyset, adLockOptimistic

        Screen.ActiveForm.ActiveControl.Clear

            With rsPRODUCT
                    Do Until .EOF
                            Screen.ActiveForm.ActiveControl.AddItem !ProductDescription
                            .MoveNext
                    Loop
            End With
  
rsPRODUCT.Close
strPRODUCT = ""

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub ProductCodeLostFocus()
On Error GoTo err

        Dim rsPRODUCT As ADODB.Recordset, strPRODUCT As String
        Set rsPRODUCT = New Recordset
        
        If banticipatedENDOWMENT = True Then
                rsPRODUCT.Open "SELECT * FROM ALISPProduct WHERE ProductDescription= '" & Screen.ActiveForm.cboProductCodeTAB1.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        Else: rsPRODUCT.Open "SELECT * FROM ALISPProduct WHERE ProductDescription= '" & Screen.ActiveForm.cboProductCode.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        End If
        
        With rsPRODUCT
                If .EOF And .BOF Then Exit Sub
                If banticipatedENDOWMENT = True Then
                        Screen.ActiveForm.cboProductCodeTAB1.Text = !ProductCode
                Else
                        Screen.ActiveForm.cboProductCode.Text = !ProductCode
                        Screen.ActiveForm.txtProductDescription = !ProductDescription
                End If
        End With

rsPRODUCT.Close

Exit Sub

err:
    ErrorMessage

End Sub

Public Sub LoadProductCode()
On Error GoTo err

        Dim rsPRODUCT As ADODB.Recordset, strPRODUCT As String
        Set rsPRODUCT = New Recordset
        
        rsPRODUCT.Open "SELECT * FROM ALISPProduct WHERE ProductCode= '" & Screen.ActiveForm.cboProductCode.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsPRODUCT
                If .EOF And .BOF Then Exit Sub
                            Screen.ActiveForm.txtProductDescription.Text = !ProductDescription
        End With
    
rsPRODUCT.Close

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub ClearListview()
On Error GoTo err

 Dim j, i As Integer
       
                j = ListView1.ListItems.Count
            
        For i = 1 To j
                ListView1.ListItems(i).Checked = False
        Next i
Exit Sub
err:
ErrorMessage
    
End Sub

Public Sub clearALLRECORD()
On Error GoTo err
Dim i
    For Each i In Screen.ActiveForm
        If TypeOf i Is TextBox Then
            i.Text = Empty
        End If
        If TypeOf i Is ComboBox Then
            i.Clear
        End If
        
            
        If TypeOf i Is DTPicker Then
                i.Value = Date
        End If
        
        If TypeOf i Is CheckBox Then
                i.Value = 0
        End If
        
        If TypeOf i Is OptionButton Then
                i.Value = 0
        End If
        
    Next i

Exit Sub

err:
    ErrorMessage
End Sub



Public Sub enableALLRECORD()
On Error GoTo err

Dim i
    For Each i In Screen.ActiveForm
            If TypeOf i Is TextBox Then
                i.Locked = False
            End If
            
            If TypeOf i Is ComboBox Then
                i.Locked = False
            End If
            
            If TypeOf i Is VScrollBar Then
                i.Enabled = True
            End If
            
            If TypeOf i Is DTPicker Then
                i.Enabled = True
            End If
            
            If TypeOf i Is UpDown Then
                i.Enabled = True
            End If
            
            If TypeOf i Is CheckBox Then
                i.Enabled = True
            End If
            
            If TypeOf i Is OptionButton Then
                i.Enabled = True
            End If
            
            If TypeOf i Is DTPicker Then
                i.Enabled = True
            End If

            

    Next i

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub cancelCMD()
        enableButtons
        clearALLRECORD
        disableALLRECORD
End Sub
Public Sub addCMD()
        clearALLRECORD
        enableALLRECORD
        disableButtons
End Sub


Public Sub disableALLRECORD()
On Error GoTo err

Dim i
    For Each i In Screen.ActiveForm
    
            If TypeOf i Is TextBox Then
                i.Locked = True
            End If
            
            If TypeOf i Is ComboBox Then
                i.Locked = True
            End If
            
            If TypeOf i Is VScrollBar Then
                i.Enabled = False
            End If
            
            If TypeOf i Is DTPicker Then
                i.Enabled = False
            End If
            
            If TypeOf i Is UpDown Then
                i.Enabled = False
            End If
            
            If TypeOf i Is CheckBox Then
                i.Enabled = False
            End If
            
            If TypeOf i Is OptionButton Then
                i.Enabled = False
            End If
            
            If TypeOf i Is DTPicker Then
                i.Enabled = False
            End If
        
    Next i
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub ClaimCodeGotFocus()
On Error GoTo err

        Dim rsCLAIMCODE As ADODB.Recordset, strCLAIMCODE As String
        Set rsCLAIMCODE = New Recordset
      
        strCLAIMCODE = "SELECT * FROM ALISPClaimCode;"
        rsCLAIMCODE.Open strCLAIMCODE, cnALIS, adOpenKeyset, adLockOptimistic

        Screen.ActiveForm.cboClaimCode.Clear

            With rsCLAIMCODE
                    Do Until .EOF
                            Screen.ActiveForm.cboClaimCode.AddItem !ClaimCodeDescription
                            .MoveNext
                    Loop
            End With
        
        Exit Sub

err:
    ErrorMessage
End Sub

Public Sub ClaimCodeLostFocus()
On Error GoTo err

        Dim rsCLAIMCODE As ADODB.Recordset
        Set rsCLAIMCODE = New Recordset
        
        rsCLAIMCODE.Open "SELECT * FROM ALISPClaimCode WHERE ClaimCodeDescription= '" & Screen.ActiveForm.cboClaimCode.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsCLAIMCODE
                If .EOF And .BOF Then Exit Sub
                            Screen.ActiveForm.cboClaimCode.Text = !ClaimCode
                            Screen.ActiveForm.txtClaimCodeDescription = !ClaimCodeDescription
        End With

rsCLAIMCODE.Close


Exit Sub

err:
    ErrorMessage
End Sub
Public Sub LoadClaimCodeLostFocus()
On Error GoTo err

        Dim rsCLAIMCODE As ADODB.Recordset
        Set rsCLAIMCODE = New Recordset
        
        rsCLAIMCODE.Open "SELECT * FROM ALISPClaimCode WHERE ClaimCode = '" & Screen.ActiveForm.cboClaimCode.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsCLAIMCODE
                If .EOF And .BOF Then Exit Sub
                        Screen.ActiveForm.txtClaimCodeDescription = !ClaimCodeDescription
        End With
rsCLAIMCODE.Close

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub ClaimStatusGotFocus()
On Error GoTo err

        Dim rsSTATUS As ADODB.Recordset
        Set rsSTATUS = New Recordset
      
        strSQL = "SELECT * FROM ALISPstatus;"
        rsSTATUS.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic

        Screen.ActiveForm.cboClaimStatus.Clear

        With rsSTATUS
                Do Until .EOF
                        Screen.ActiveForm.cboClaimStatus.AddItem !StatusDescription
                        .MoveNext
                Loop
        End With
        
Exit Sub

rsSTATUS.Close
strSQL = ""

err:
    ErrorMessage
End Sub


Public Sub ClaimStatusLostFocus()
On Error GoTo err

        Dim rsSTATUS As ADODB.Recordset
        Set rsSTATUS = New Recordset
        
        rsSTATUS.Open "SELECT * FROM ALISPStatus WHERE StatusDescription= '" & Screen.ActiveForm.cboClaimStatus.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsSTATUS
                If .EOF And .BOF Then Exit Sub
                            Screen.ActiveForm.cboClaimStatus.Text = !StatusCode
                            Screen.ActiveForm.txtStatusDesc = !StatusDescription
        End With

Exit Sub

rsSTATUS.Close

err:
    ErrorMessage
End Sub

Public Sub LoadClaimStatus()
On Error GoTo err

        Dim rsSTATUS As ADODB.Recordset
        Set rsSTATUS = New Recordset
        
        rsSTATUS.Open "SELECT * FROM ALISPStatus WHERE StatusCode = '" & Screen.ActiveForm.cboClaimStatus.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsSTATUS
                If .EOF And .BOF Then Exit Sub
                            Screen.ActiveForm.txtStatusDesc = !StatusDescription
        End With

Exit Sub

rsSTATUS.Close

err:
    ErrorMessage
End Sub

