VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASTerminationLandlord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub LoadDefaultDetails()
On Error GoTo err
    With frmODASMSiteMaintanance
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPPlotSite PS, ODASMJobBriefItems JBI WHERE PS.SiteNo = '" & .txtSiteNo.Text & "' and PS.JobBriefItemNo = JBI.JobBriefItemNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
        .txtExpiryDate.Text = rsFindRecord!expirydate
        .txtCommencementDate.Text = rsFindRecord!CommencementDate
        .txtLeaseDuration.Text = rsFindRecord!Duration
        Set rsFindRecord = Nothing
    End With
Exit Sub
err:
ErrorMessage
End Sub
Public Sub loadDEFAULTS()

On Error GoTo err
    With Screen.ActiveForm
            
            .txtNoticeDate.Text = Date
            .txtTerminatedBy.Text = CurrentUserName
        If SchedulingMain.txtTask.Text = "N6" Then
            .txtRecoveryRatio.Text = 0
            .txtTerminationDate.Text = Date
        ElseIf SchedulingMain.txtTask.Text = "N12" Then
            .txtRenewalDate.Text = Date
        End If
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadRECORD()
On Error GoTo err
    With Screen.ActiveForm
            
            Set rsCONTROL = New ADODB.Recordset
            
            strSQL = "Select (ODASMLeaseAgreement.CommencementDate) As ComDate,(ODASMLeaseAgreement.ExpiryDate) as ExDate,ODASMLeaseAgreement.* ,ODASPPlot.*, ODASPAccount.* from ODASMLeaseAgreement, ODASPPlot, ODASPAccount Where ODASMLeaseAgreement.ContractNo =  '" & Screen.ActiveForm.txtContractNo.Text & "' and ODASMLeaseAgreement.PlotNo = ODASPPlot.PlotNo and ODASMLeaseAgreement.AccountNo = ODASPAccount.AccountNo "
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
            If rsCONTROL.EOF Or rsCONTROL.EOF Then Exit Sub
            .txtAgreementDate.Text = rsCONTROL!AgreementDate
            .txtLeaseDuration.Text = rsCONTROL!LeaseDuration & ""
            .txtCommencementDate.Text = rsCONTROL!ComDate
            .txtExpiryDate.Text = rsCONTROL!ExDate
            .txtLandLord.Text = rsCONTROL!AccountNo
'            .txtNarration.Text = rsCONTROL!Narration & ""
            .txtSignedBy.Text = rsCONTROL!SignedBy & ""
          
            If IsNull(rsCONTROL!NoticeDate) = True Then
                    .DTPickerNoticeDate.Value = Date
                    .txtNoticeDate.Text = Date
            Else
                    .txtNoticeDate.Text = rsCONTROL!NoticeDate
                    .DTPickerNoticeDate.Value = rsCONTROL!NoticeDate
            End If
            
            .txtPhysicalLocation.Text = rsCONTROL!PhysicalLocation
            .txtPlotName.Text = rsCONTROL!PlotName
            .txtPlotNo.Text = rsCONTROL!PLotNo
            
            If IsNull(rsCONTROL!TerminatedBy) = True Then
                    .txtTerminatedBy.Text = CurrentUserName
            Else
                    .txtTerminatedBy.Text = rsCONTROL!TerminatedBy
            End If
            
            If SchedulingMain.txtTask.Text = "N6" Then

             .txtRentPaid.Text = rsCONTROL!RentPaid & ""
             .txtRentRecovered.Text = rsCONTROL!RentRecovered & ""
             .txtRecoveryRatio.Text = rsCONTROL!RecoveryRatio & ""
            
                     If IsNull(rsCONTROL!TerminationDate) = True Then
                             .DTPickerTerminationDate.Value = Date
                             .txtTerminationDate.Text = Date
                     Else
                             .txtTerminationDate.Text = rsCONTROL!TerminationDate
                             .DTPickerTerminationDate.Value = rsCONTROL!TerminationDate
                    End If
            ElseIf SchedulingMain.txtTask.Text = "N12" Then
                    .txtTerminatedBy.Text = CurrentUserName
                    .txtRenewalDate.Text = Date
                    
                Set rsFIND = New ADODB.Recordset: Set rsFindRecord = New ADODB.Recordset
                
                rsFindRecord.Open "Select * From ODASPPlotMast Where ContractNo = '" & Screen.ActiveForm.txtContractNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsFindRecord!Renewal = "Y" Then
                     rsFIND.Open "Select * From ODASMLeaseRenewals Where ContractNo = '" & Screen.ActiveForm.txtContractNo.Text & "' and ReNewalNo = '1';", cnCOMMON, adOpenKeyset, adLockOptimistic
                ElseIf IsNull(rsFindRecord!Renewal) Then Exit Sub
                Else
                     rsFIND.Open "Select * From ODASMLeaseRenewals Where ContractNo = '" & Screen.ActiveForm.txtContractNo.Text & "' and ReNewalNo = '" & rsCONTROL!Renewal & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                End If
                     frmODASMLeaseRenewal.txtMonth.Text = rsFIND!RenewalPeriod
                     frmODASMLeaseRenewal.txtRenewalEndDate.Text = rsFIND!RenewalExpiryDate
            
            End If
 
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableRECORD()
On Error GoTo err
        With Screen.ActiveForm
            .txtNarration.Locked = False
            .txtNoticeDate.Locked = False
            .DTPickerNoticeDate.Enabled = True
            If SchedulingMain.txtTask.Text = "N1" Then
            .Frame2.Enabled = True
            .optRenewal.Enabled = True
            .optTermination.Enabled = True
            .txtMonth.Locked = False
            .txtYear.Locked = False
            .txtMonth.Enabled = True
            .txtYear.Enabled = True
            .chkAll.Enabled = True
            .UpDownRenewalDuration.Enabled = True
            ElseIf SchedulingMain.txtTask.Text = "N6" Then
                .txtRecoveryRatio.Locked = False
                .txtRentRecovered.Locked = False
                .txtTerminationDate.Locked = False
                .DTPickerTerminationDate.Enabled = True
                .DTPickerNoticeDate.Enabled = True
                .cboTerminationCode.Locked = False
            End If
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub GoodData()
        With frmODASMSiteMaintanance
                bSaveRECORD = False
                
                If .optYES.Value = True And .cboStaff.Text = Empty Then
                        MsgBox "Enter  the Staff who did the Maintanace"
                        .txtDoneBy.SetFocus

                Else
                        bSaveRECORD = True
                End If
        End With
End Sub
Public Sub ValidateRECORD()
On Error GoTo err
        With Screen.ActiveForm
        
                bSaveRECORD = False
                
                If .txtContractNo.Text = Empty Then
                        MsgBox "The Contract No Cannot be Left Blank"
                        .txtContractNo.SetFocus
'                ElseIf .txtLandLordWitness.Text = Empty Then
'                        MsgBox "The Narration Cannot be Blank"
'                        .txtLandLordWitness.SetFocus
                ElseIf .txtNarration.Text = Empty Then
                        MsgBox "The Narration Cannot be Blank"
                        .txtNarration.SetFocus
                ElseIf .txtNoticeDate.Text = Empty Then
                        MsgBox "The Notice Date Cannot be Blank"
                        .txtNoticeDate.SetFocus
                        
                ElseIf .txtTerminatedBy.Text = Empty Then
                        MsgBox "Termination By Cannot Be Blank"
                        .txtTerminatedBy.SetFocus
                If SchedulingMain.txtTask.Text = "N6" Then
                    If .txtYears.Text = Empty Then
                        MsgBox "Specify the Renewal Period in Years"
                        .txtYears.SetFocus
                    ElseIf .txtMonths.Text = Empty Then
                        MsgBox "Specify the addition months"
                        .txtMonths.SetFocus
                    
                ElseIf SchedulingMain.txtTask.Text = "N6" Then
                    If .txtTerminationDate.Text = Empty Then
                            MsgBox "The Termination Date Cannot Blank"
                            .txtTerminationDate.SetFocus
                    
                    ElseIf .cboTerminationCode.Text = Empty Then
                        MsgBox "The Termination Code is Required........."
                        .cboTerminationCode.SetFocus
                    Else
                    bSaveRECORD = True

                    End If
                End If
                End If
                Else
                        bSaveRECORD = True
                End If
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub saveMaintanance()
    With frmODASMSiteMaintanance
        
        Set rsSAVE = New ADODB.Recordset
        rsSAVE.Open "SELECT * FROM ODASMMaintenance WHERE MaintenanceNo = '" & .txtMaintananceNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsSAVE.EOF And rsSAVE.BOF Then Exit Sub
            If .optYES.Value = True Then
                rsSAVE!Maintained = "Y"
                rsSAVE!Comments = .txtNarration.Text
                Else
                rsSAVE!Maintained = "N"
                rsSAVE!Comments = .txtNarration.Text
            End If
            
            rsSAVE.Update
            Set rsSAVE = Nothing
            UpdateJobBriefItems
    End With
End Sub
Public Sub UpdateJobBriefItems()
On Error GoTo err
With frmODASMSiteMaintanance
    Set rsSAVE = New ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT Min(ODASMMaintenance.MaintenanceDate)as Next,ODASMMaintenance.SiteNo FROM ODASMMaintenance WHERE SiteNo = '" & .txtSiteNo.Text & "' and MaintenanceDate>'" & Format(Date, "MMMM dd,yyyy") & "' and (Maintained ='N' or Maintained is null) GROUP By SiteNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
    
    rsSAVE.Open "SELECT * FROM ODASMJobBriefItems JBI, ODASPPlotSite PS WHERE JBI.JobBriefItemNo = PS.JobBriefItemNo and PS.SiteNo= '" & .txtSiteNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsSAVE.EOF And rsSAVE.BOF Then Exit Sub
        rsSAVE!MaintananceDueDate = rsFindRecord!Next
        rsSAVE.Update
End With
Exit Sub
err:
ErrorMessage
End Sub
Public Sub saveRecord()
On Error GoTo err

    With Screen.ActiveForm

            Set rsSAVE = New ADODB.Recordset
            
            strSQL = "Select * from ODASMLeaseAgreement Where ContractNo = '" & .txtContractNo.Text & "' "
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.BOF Or rsSAVE.EOF Then Exit Sub
                If SchedulingMain.txtTask.Text = "N6" Then
                                rsSAVE!TerminationDate = Format(Date, "MMMM dd,yyyy")
                                rsSAVE!TerminationCode = .cboTerminationCode.Text
                                rsSAVE!TerminatedBy = "xxx"
                                rsSAVE!Terminated = "Y"
                                rsSAVE!RecoveryRatio = CDbl(.txtRecoveryRatio.Text)
                                rsSAVE!RentRecovered = CDbl(.txtRentRecovered.Text) & ""
                                rsSAVE!Narration = "xxx"
                End If
                
                If SchedulingMain.txtTask.Text = "N12" Then
                    
                    Dim rsRenew As ADODB.Recordset
                    Set rsRenew = New ADODB.Recordset
                    
                    rsRenew.Open "Select * From ODASMLeaseRenewals WHERE RenewalNo = '" & rsSAVE!Renewal & "' and ContractNo = '" & .txtContractNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    If rsRenew.RecordCount = 0 Then
                        rsRenew.AddNew
                        rsRenew!RenewalNo = rsSAVE!Renewal
                        rsRenew!ContractNo = rsSAVE!ContractNo
                    End If
                        rsRenew!RenewalDate = Format(.txtRenewalDate, "MMMM dd,yyyy")
                        rsRenew!RenewalStartDate = rsSAVE!expirydate
                            If .txtYear = Empty Then
                            rsRenew!RenewalPeriod = .txtMonth.Text
                            Else
                            rsRenew!RenewalPeriod = (.txtYear.Text * 12) + (.txtMonth.Text)
                            End If
                        rsRenew!RenewalPeriod = DateDiff("m", Format(.txtRenewalDate.Text, "MMMM dd,yyyy"), Format(.txtRenewalEndDate.Text, "MMMM dd,yyyy"))
                        rsRenew!RenewalExpiryDate = Format(frmODASMLeaseRenewal.txtRenewalEndDate.Text, "MMMM dd,yyyy")
                        
                        rsRenew.Update
                        
                        rsSAVE!Renewal = rsSAVE!Renewal + 1
                        rsSAVE!Renewed = "N"
                    
                ElseIf SchedulingMain.txtTask.Text = "N1" Then
                    rsSAVE!NoticePrepared = "Y"
                    rsSAVE!NoticePreparedBy = "xxx"
                    rsSAVE!NoticeApprovedBy = "xxx"
                    rsSAVE!NoticeApprovalDate = Date
                    rsSAVE!NoticeApproved = "Y"
                    rsSAVE!NoticeAuthorized = "N"
                    rsSAVE!ReasonsForNotice = "xxx"
                    rsSAVE!NoticeDate = Format(.txtNoticeDate.Text, "MMMM dd,yyyy")
                    rsSAVE!ActionBy = "xxx"
                    rsSAVE!ActionDate = Date
                    
                     rsSAVE.Update
                    rsSAVE.Requery
                    If frmODASMPrepareNotice.optRenewal.Value = True Then
                        
                    Else:
                      End If
                    
                    bSaveRECORD = False
                    
                    
                End If
           
     
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
Public Sub UpdateLease()
On Error GoTo err

    With frmODASMLeaseAgreement

            Set rsSAVE = New ADODB.Recordset
            
            strSQL = "Select * from ODASMLeaseAgreement Where ContractNo = '" & .txtContractNo.Text & "' "
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.BOF Or rsSAVE.EOF Then Exit Sub
                               
            rsSAVE!TerminationDate = frmODASMLeaseAgreement.txtNoticeDate.Text
            rsSAVE!TerminatedBy = CurrentUserName
            rsSAVE!Terminated = "Y"
            rsSAVE!TerminationCode = frmODASMLeaseAgreement.cboTerminationCode.Text
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

Public Sub updatePlotMastRenewal()
On Error GoTo err
    With frmODASMPrepareNotice
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPPlotMast WHERE ContractNo = '" & frmODASMPrepareNotice.txtContractNo.Text & "' and renewal is not null;", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
        rsFindRecord.MoveFirst
        Do While Not rsFindRecord.EOF
            rsFindRecord!expirydate = Format(.txtRenewalEndDate.Text, "MMMM dd,yyyy")
            If rsFindRecord!Renewal = "N" Then
                rsFindRecord!Renewal = "N"
            Else
                rsFindRecord!Renewal = rsFindRecord!Renewal + 1
            End If
            
            rsFindRecord.Update
        rsFindRecord.MoveNext
        Loop
    End With

Exit Sub
err:
ErrorMessage
End Sub
Public Sub updatePlotSiteRenewal()
On Error GoTo err
    With frmODASMLeaseRenewal
        Dim SDate, EDate As Variant
        
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT  ODASPPlotSite.MastNo,ODASPPlotSite.SiteNo, ODASPPlotMast.* FROM ODASPPlotSite, ODASPPlotMast WHERE ODASPPlotMast.ContractNo = '" & .txtContractNo.Text & "' and ODASPPlotMast.MastNo = ODASPPlotSite.MastNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
        rsFindRecord.MoveFirst
        
        Do While Not rsFindRecord.EOF
            SDate = DateAdd("d", 1, rsFindRecord!expirydate)
            rsFindRecord!expirydate = Format(.txtRenewalEndDate.Text, "MMMM dd,yyyy")
            
            rsFindRecord.Update
            
                EDate = rsFindRecord!expirydate
                
                .ProgressBar2.Visible = True
                '.ProgressBar2.Max = DateDiff("d", SDate, EDate)
                '.ProgressBar2.Min = 0
                
                Do While SDate <= EDate
                    Set rsSiteSchedule = New ADODB.Recordset
                    rsSiteSchedule.Open "SELECT * FROM ODASMSiteSchedule WHERE SiteNo = '" & rsFindRecord!SiteNo & "' and ScheduleDate = '" & Format(SDate, "MMMM dd,yyyy") & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    If rsSiteSchedule.RecordCount = 0 Then
                        rsSiteSchedule.AddNew
                        rsSiteSchedule!SiteNo = rsFindRecord!SiteNo
                        rsSiteSchedule!ScheduleDate = SDate
                        rsSiteSchedule!Reserved = "N"
                        rsSiteSchedule!Allocated = "N"
                    End If
                    rsSiteSchedule.Update
                    SDate = DateAdd("d", 1, SDate)
                   .ProgressBar2.Value = .ProgressBar2.Value + 1
                Loop
        .ProgressBar2.Value = 0
        rsFindRecord.MoveNext
        Loop
    End With
Exit Sub
err:
ErrorMessage
End Sub
Public Sub updateSiteSchedule()
On Error GoTo err
    
Exit Sub
err:
ErrorMessage
End Sub
Public Sub updatePlotRenewal()
On Error GoTo err

    With frmODASMLeaseRenewal
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPPlot WHERE PlotNo = '" & .txtPlotNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
            rsFindRecord!CommencementDate = Format(.txtRenewalDate.Text, "MMMM dd,yyyy")
            rsFindRecord!expirydate = Format(.txtRenewalEndDate.Text, "MMMM dd,yyyy")
            rsFindRecord!LeaseDuration = .txtYear.Text
            rsFindRecord.Update
            rsFindRecord.Requery
    End With
Exit Sub
err:

End Sub
Public Sub updateContractTermination()
On Error GoTo err
    With frmODASMPrepareNotice
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASMLeaseAgreement WHERE ContractNo = '" & frmODASMPrepareNotice.txtContractNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
        rsFindRecord.MoveFirst
        Do While Not rsFindRecord.EOF
            rsFindRecord!TerminationDate = Date
            rsFindRecord!TerminatedBy = CurrentUserName
            rsFindRecord!Terminated = "Y"
            
            rsFindRecord.Update
        rsFindRecord.MoveNext
        Loop
    End With

Exit Sub
err:
ErrorMessage
End Sub
Public Sub updateRenewal()
On Error GoTo err

    With frmODASMLeaseRenewal
       Dim rsRenew As ADODB.Recordset
       Set rsRenew = New ADODB.Recordset
                    
           rsRenew.Open "Select * From ODASMLeaseRenewals WHERE RenewalNo = '" & rsSAVE!Renewal & "' and ContractNo = '" & .txtContractNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
           If rsRenew.RecordCount = 0 Then
            rsRenew.AddNew
            rsRenew!RenewalNo = rsSAVE!Renewal
            rsRenew!ContractNo = rsSAVE!ContractNo
           End If
            rsRenew!RenewalStartDate = DateAdd("d", 1, rsSAVE!expirydate)
           If frmODASMPrepareNotice.optRenewal.Value = True Then
             rsRenew!RenewalPeriod = DateDiff("m", Format(.txtExpiryDate.Text, "MMMM dd,yyyy"), Format(.txtRenewalEndDate.Text, "MMMM dd,yyyy"))
             rsRenew!RenewalExpiryDate = Format(.txtRenewalEndDate.Text, "MMMM dd,yyyy")
           End If
             rsRenew!NoOfBillBoards = K
             rsRenew.Update
        
    End With
Exit Sub
err:

End Sub
