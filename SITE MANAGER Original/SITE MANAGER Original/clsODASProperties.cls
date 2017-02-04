VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub loadDEFAULTS()
On Error GoTo err
    With frmODASMSiteRegistration
        .txtPropertyAmountDue.Text = 0
        .txtPropertyCommencementDate.Text = Date
        .DTPickerCommencementDate.Value = Date
        .txtPropertyDateAssigned.Text = Date
        .DTPickerPropertyDateAssigned.Value = Date
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub clearRECORD()
On Error GoTo err
    With frmODASMSiteRegistration
        .txtPropertyAmountDue.Text = 0
        .txtPropertyCode.Text = Empty
        .txtPropertyDateAssigned.Text = Date
        .txtPropertyOtherDetails.Text = Empty
        .txtPropertyCommencementDate = Date
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableRECORD()
On Error GoTo err
    With frmODASMSiteRegistration
        .txtPropertyAmountDue.Locked = False
        .txtPropertyCode.Locked = False
        .txtPropertyOtherDetails.Locked = False
        .DTPickerPropertyCommencementDate.Enabled = True
        .DTPickerPropertyDateAssigned.Enabled = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadRECORD()
On Error GoTo err
    With frmODASMSiteRegistration
        
            Set rsCONTROL = New ADODB.Recordset
            rsCONTROL.Open "SELECT * FROM ODASMSiteProperties  WHERE SiteNo = '" & .txtSiteNo.Text & "' and PropertyCode = '" & .txtPropertyCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
            If IsNull(rsCONTROL!AmountDue) = True Then
                .txtPropertyAmountDue.Text = 0
            Else: .txtPropertyAmountDue.Text = FormatNumber(rsCONTROL!AmountDue, 2)
            End If
            
            If IsDate(rsCONTROL!CommencementDate) = True Then .txtPropertyCommencementDate.Text = rsCONTROL!CommencementDate & ""
            If IsDate(rsCONTROL!DateAssigned) = True Then .txtPropertyDateAssigned.Text = rsCONTROL!DateAssigned & ""
            .txtPropertyOtherDetails.Text = rsCONTROL!OtherDetails & ""
            .txtPropertyTransactionNo.Text = rsCONTROL!TransactionNo
            
            Set rsCONTROL = Nothing
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub saveRecord()
On Error GoTo err
    With frmODASMSiteRegistration

        j = .ListALLProperties.ListItems.Count
        For i = 1 To j
            If .ListALLProperties.ListItems(i).Checked = True Then
            
                .txtPropertyCode = Trim(.ListALLProperties.ListItems(i).Text)
                
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open "SELECT * FROM ODASMSiteProperties  WHERE SiteNo = '" & Trim(.txtSiteNo.Text) & "' and PropertyCode = '" & .txtPropertyCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
                If rsSAVE.EOF And rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!SiteNo = .txtSiteNo.Text
                        rsSAVE!PropertyCode = .txtPropertyCode.Text
                        rsSAVE!Status = "ACTIVE"
                        rsSAVE!PreparedBY = InUserName
                        rsSAVE!DatePrepared = Date
                End If
                
                rsSAVE!OtherDetails = .txtPropertyOtherDetails
                rsSAVE!AmountDue = CCur(.txtPropertyAmountDue.Text)
                rsSAVE!DateAssigned = .txtPropertyDateAssigned.Text
                rsSAVE!CommencementDate = .txtPropertyCommencementDate.Text
                rsSAVE!PLotNo = .txtPlotNo.Text
                
                rsSAVE.Update
                
            End If
         Next i

        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub deleteRECORD()
On Error GoTo err
With frmODASMSiteRegistration
        Set rsSAVE = New ADODB.Recordset
        rsSAVE.Open "Delete FROM ODASMSiteProperties WHERE SiteNo = '" & Trim(.txtSiteNo.Text) & "' and PropertyCode = '" & .txtPropertyCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub updateRECORD()
       
    If ValidRecord = True Then
        saveRecord
        disableALLRECORD
    End If
    showALLPROPERTIES
    showACTUALPROPERTIES

End Sub

Public Function ValidRecord()
On Error GoTo err
    
    With frmODASMSiteRegistration
            ValidRecord = False
            
            If .txtSiteNo.Text <= "" Then
                strMessage = "The Site Number is required ..........."
                .txtSiteNo.SetFocus
            
            ElseIf .txtPropertyDateAssigned.Text = Empty Then
                strMessage = "The Date Assigned cannot not be ignored"
                .txtPropertyDateAssigned.SetFocus
            
            ElseIf k = 0 Then
                strMessage = "Please choose one or more properties Required ..........."
                .ListView1.SetFocus

            Else
                ValidRecord = True
            End If
            
            If Not ValidRecord Then
                MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
            End If
            
    End With
Exit Function

err:
    ErrorMessage
End Function

