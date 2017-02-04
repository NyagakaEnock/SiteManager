VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub loadDEFAULTS()

On Error GoTo err
    With frmODASMSendNotice
               .txtDispatchDate.Text = Date
               .txtDispatchedBy.Text = CurrentUserName
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadTownPlot()
With Screen.ActiveForm
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM ODASPTown, ODASPPlot WHERE ODASPTown.TownCode = '" & .txtTownCode & "' and ODASPTown.TownCode = ODASPPlot.TownCode;", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
    .txtTown.Text = rsFindRecord!Town
    .txtPlotName.Text = rsFindRecord!PlotName
    .txtPlotNo.Text = rsFindRecord!PLotNo
    .txtPhysicalLocation.Text = rsFindRecord!PhysicalLocation
    
End With
End Sub
Public Sub loadRECORD()
On Error GoTo err
    With frmODASMSendNotice
            
            Set rsCONTROL = New ADODB.Recordset
            
            strSQL = "Select * from ODASMLeaseAgreement, ODASPPlot, ODASPAccount Where ODASMLeaseAgreement.ContractNo =  '" & frmODASMSendNotice.txtContractNo.Text & "' and ODASMLeaseAgreement.PlotNo = ODASPPlot.PlotNo and ODASMLeaseAgreement.AccountNo = ODASPAccount.AccountNo "
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
            If rsCONTROL.EOF Or rsCONTROL.EOF Then Exit Sub
                
            .txtLeaseDuration.Text = rsCONTROL!LeaseDuration & ""
            .txtCommencementDate.Text = rsCONTROL!CommencementDate
            .txtExpiryDate.Text = rsCONTROL!expirydate
            .txtLandLord.Text = rsCONTROL!AccountNo
            
            
            .txtPhysicalLocation.Text = rsCONTROL!PhysicalLocation
            .txtPlotName.Text = rsCONTROL!PlotName
            .txtPlotNo.Text = rsCONTROL!PLotNo
            
            

    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadDISPATCHDETAILS()
On Error GoTo err
        With frmODASMSendNotice
            
            Set rsCONTROL = New ADODB.Recordset
            
            strSQL = "Select * from ODASMDispatch Where ODASMDispatch.DispatchReferenceNo =  '" & frmODASMSendNotice.txtDispatchReferenceNo.Text & "' "
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
            If IsNull(rsCONTROL!DispatchedBy) = True Then
                    .txtDispatchedBy.Text = CurrentUserName
            Else
                    .txtDispatchedBy.Text = rsCONTROL!DispatchedBy
            End If
            
            If IsNull(rsCONTROL!DispatchDate) = True Then
                    .txtDispatchDate.Text = Date
            Else
                    .txtDispatchDate.Text = rsCONTROL!DispatchDate
            End If

                    
        End With
Exit Sub
err:
    ErrorMessage
End Sub



Public Sub ValidateRECORD()
On Error GoTo err
        With frmODASMSendNotice
        
                bSaveRECORD = False
                
                If .txtContractNo.Text = Empty Then
                        MsgBox "The Contract No Cannot be Left Blank......."
                        .txtContractNo.SetFocus
                        
                ElseIf .txtDispatchedBy.Text = Empty Then
                        MsgBox "The Dispatched Cannot be Blank............."
                        .txtDispatchedBy.SetFocus
                        
                ElseIf .txtDispatchDate.Text = Empty Then
                        MsgBox "The Termination Code is Required........."
                        .txtDispatchDate.SetFocus
                Else
                        bSaveRECORD = True
                End If

        End With
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub generateReferenceNO()
On Error GoTo err
        Set rsCONTROL = New Recordset
        
        strSQL = "Select * from ODASPLAstNumbers Where AutoDispatchReferenceNo = 'Y';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
        
                If .EOF Or .BOF = True Then Exit Sub
                
                frmODASMSendNotice.txtDispatchReferenceNo.Text = !DispatchReferenceNo
                
                Select Case Len(frmODASMSendNotice.txtDispatchReferenceNo.Text)
                        Case 1: frmODASMSendNotice.txtDispatchReferenceNo.Text = Trim(!RequisitionPrefix) + "0000" + Trim(frmODASMSendNotice.txtDispatchReferenceNo.Text)
                        Case 2: frmODASMSendNotice.txtDispatchReferenceNo.Text = Trim(!RequisitionPrefix) + "000" + Trim(frmODASMSendNotice.txtDispatchReferenceNo.Text)
                        Case 3: frmODASMSendNotice.txtDispatchReferenceNo.Text = Trim(!RequisitionPrefix) + "00" + Trim(frmODASMSendNotice.txtDispatchReferenceNo.Text)
                        Case 4: frmODASMSendNotice.txtDispatchReferenceNo.Text = Trim(!RequisitionPrefix) + "0" + Trim(frmODASMSendNotice.txtDispatchReferenceNo.Text)
                        Case 5: frmODASMSendNotice.txtDispatchReferenceNo.Text = Trim(!RequisitionPrefix) + Trim(frmODASMSendNotice.txtDispatchReferenceNo.Text)
                End Select
                                
                !DispatchReferenceNo = Str(Val(!DispatchReferenceNo) + 1)
                .Update
                .Requery
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub saveRecord()
On Error GoTo err

    With frmODASMSendNotice

            Set rsSAVE = New ADODB.Recordset
            
            strSQL = "Select * from ODASMLeaseAgreement Where ContractNo = '" & .txtContractNo.Text & "' "
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.BOF Or rsSAVE.EOF Then Exit Sub
                               
            rsSAVE!DispatchReferenceNo = .txtDispatchReferenceNo
            rsSAVE!NoticeDispatched = "Y"
            bSaveRECORD = False

            rsSAVE.Update
            rsSAVE.Requery
  End With

Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If

End Sub

Public Sub SaveDispatch()
On Error GoTo err

    With frmODASMSendNotice

            Set rsSAVE = New ADODB.Recordset
            
            strSQL = "Select * from ODASMDispatch Where DispatchReferenceNo = '" & .txtDispatchReferenceNo.Text & "' "
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.BOF Or rsSAVE.EOF Then
                    generateReferenceNO
                    rsSAVE.AddNew
                    rsSAVE!DispatchReferenceNo = .txtDispatchReferenceNo
                    rsSAVE!DispatchedBy = .txtDispatchedBy.Text
                    rsSAVE!DispatchDate = .txtDispatchDate.Text
            End If
            
            rsSAVE!ContractNo = .txtContractNo.Text
            rsSAVE!DocumentType = "SENDNOTICE"
            bSaveRECORD = False

            rsSAVE.Update
            rsSAVE.Requery
  End With

Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If

End Sub




