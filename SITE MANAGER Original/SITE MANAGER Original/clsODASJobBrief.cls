VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASJobBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rsTASK As ADODB.Recordset, strTASK As String

Public Sub generateTASKS()
On Error GoTo err
        
        Dim rsTASK As ADODB.Recordset, strTASK As String
        Set rsTASK = New Recordset
        
        strTASK = "SELECT * FROM ODASMJobBriefItems WHERE JobBriefNo = '" & CurrentRecord & "'"
        rsTASK.Open strTASK, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsTASK.EOF Or rsTASK.BOF Then Exit Sub
        
        Do While Not rsTASK.EOF
                
                Dim rsDEPT As ADODB.Recordset, strDEPT As String
                Set rsDEPT = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPMediaTask WHERE MediaCode = '" & rsTASK!MediaCode & "'"
                rsDEPT.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                If rsDEPT.EOF Or rsDEPT.BOF Then

                Else
                        Do While Not rsDEPT.EOF
                            
                                Set rsSAVE = New Recordset
                                strSQL = "SELECT * FROM ODASMJobCARD WHERE JobCardNo = '" & CurrentRecord & "' and DepartmentCode = '" & rsDEPT!DepartmentCode & "'"
                                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                                If rsSAVE.EOF Or rsSAVE.BOF Then
                                        rsSAVE.AddNew
                                        rsSAVE!JobCardNo = CurrentRecord
                                        rsSAVE!DepartmentCode = rsDEPT!DepartmentCode
                                        rsSAVE!DatePrepared = Date
                                        rsSAVE!PreparedBY = CurrentUserName
                                        rsSAVE!Status = "START"
                                        rsSAVE!Prepared = "Y"
                                        rsSAVE.Update
                                End If
                            
                            rsDEPT.MoveNext
                        Loop
                End If
                
                rsTASK.MoveNext
        Loop

Exit Sub

err:
    ErrorMessage
End Sub



Public Sub openJOBBRIEF()
On Error GoTo err
        
        Dim rsOPEN As ADODB.Recordset, strOPEN As String
        
        Set rsCHECK = New Recordset
        
        strCHECK = "SELECT * FROM ODASMJobCard WHERE JobCardNo = '" & CurrentRecord & "' and Opened ='N' "
        rsCHECK.Open strCHECK, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCHECK.EOF Or rsCHECK.BOF Then
                
                '/update the Job Brief File to Indicate that the
                '/ Job Cards are open
                
                Set rsOPEN = New Recordset
                
                strOPEN = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & CurrentRecord & "'"
                rsOPEN.Open strOPEN, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                If rsOPEN.EOF Or rsOPEN.BOF Then Exit Sub
                    rsOPEN!Status = "OPENED"
                    rsOPEN!opened = "Y"
                    rsOPEN!Closed = "N"
                    rsOPEN!DateOpened = Date
                    rsOPEN!openedBy = CurrentUserName
                    rsOPEN.Update
                    rsOPEN.Requery
        End If
        
strOPEN = Empty

Exit Sub

err:
    ErrorMessage
End Sub





Public Sub startJOBBRIEF()
On Error GoTo err
        
        Dim rsSTART As ADODB.Recordset, strSTART As String
        Set rsSTART = New ADODB.Recordset
        
        strSTART = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & CurrentRecord & "'"
        rsSTART.Open strSTART, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsSTART.EOF Or rsSTART.BOF Then Exit Sub
        
        If rsSTART.EOF Or rsSTART.BOF Then Exit Sub
        rsSTART!AuthorizedBy = CurrentUserName
        rsSTART!DateAuthorized = Date
        rsSTART!Authorized = "Y"
        rsSTART!Status = "STARTED"
        rsSTART!Started = "Y"
        rsSTART!DateStarted = Date
        rsSTART!StartedBy = CurrentUserName
        rsSTART.Update
        rsSTART.Requery
        

'rsSTART.Close

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub printDataBaseReport()
        Set rsPrint = New Recordset
        
        strSQL = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & CurrentRecord & "'"
        rsPrint.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsPrint!Closed <> "Y" Then Exit Sub
        Load frmODASROutDoorDataBase
        frmODASROutDoorDataBase.Show 1, SchedulingMain
       
End Sub

Public Sub closeJOBBRIEF()
On Error GoTo err
        
        Set rsSAVE = New Recordset
        
        strSQL = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & CurrentRecord & "'"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub

        Set rsCONTROL = New Recordset
        
        strCONTROL = "SELECT * FROM ODASMJobCard WHERE JobCardNo = '" & CurrentRecord & "' and Closed = 'N'"
        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF And rsCONTROL.BOF Then
                
                        rsSAVE!Status = "JOB BRIEF-CLOSED"
                        rsSAVE!Closed = "Y"
                        rsSAVE!DateClosed = Date
                        rsSAVE!ClosedBy = CurrentUserName
                        rsSAVE.Update
                        rsSAVE.Requery
                        UpdateJobBriefItems
                        updateJobBriefDates
                        printDataBaseReport
        End If

rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateJobBriefDates()
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT min(CommencementDate) as startDate,max(ExpiryDate)as Enddate,JobBriefNo FROM ODASMJobBriefItems WHERE JobBriefNo = '" & CurrentRecord & "' GROUP BY JobBriefNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsSAVE = New ADODB.Recordset
    rsSAVE.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & CurrentRecord & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsSAVE.EOF And rsSAVE.BOF Then Exit Sub
    
    rsSAVE!CommencementDate = rsFindRecord!StartDate
    rsSAVE!ExpiryDate = rsFindRecord!EndDate
    rsSAVE.Update
    
End Sub
Public Sub UpdateJobBriefItems()
On Error GoTo err
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT JB.*,(JBI.CommencementDate) as ComDate, (JBI.ExpiryDate) as ExDate, JBI.*,PS.*,P.* FROM ODASMJobBrief JB,ODASMJobBriefItems JBI,ODASPPlotSite PS,ODASPPlot P WHERE JBI.JobBriefNo ='" & CurrentRecord & "' and JB.JobBriefNo = JBI.JobBriefNo and (JBI.SiteNo <> '" & "N/A" & "' and JBI.SiteNo is not null) and JBI.SiteNo = PS.SiteNo and PS.PlotNo = P.PlotNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
    num = rsFindRecord.RecordCount
    
    rsFindRecord.MoveFirst
    Do While Not rsFindRecord.EOF
        
        Set rsfindrecord3 = New ADODB.Recordset
        rsfindrecord3.Open "SELECT * FROM ODASMCouncilRatesPayable WHERE JobBriefItemNo = '" & rsFindRecord!JobBriefItemNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        Dim rate As String
        If rsfindrecord3.EOF And rsfindrecord3.BOF Then
            rate = 0
            Else
            rate = rsfindrecord3!RatePayable
        End If
                
                If rsFindRecord!MaintananceInterval = "N" Or IsNull(rsFindRecord!MaintananceInterval) Then GoTo Continue
                
                Dim MaintenanceNo, Times, Period As String
                Set rsFindRecord2 = New ADODB.Recordset
                
                rsFindRecord2.Open "SELECT * FROM ODASPInterval WHERE IntervalCode = '" & rsFindRecord!MaintananceInterval & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsFindRecord2.EOF And rsFindRecord2.BOF Then GoTo Continue
                rsFindRecord!MaintananceDueDate = DateAdd("M", rsFindRecord2!NoOfMonths, Date)
                
                If rsFindRecord2!NoOfMonths = 0 Then GoTo Continue
                Times = 1: Period = 0
                
                frmODASMOperation.ProgressBar1.Visible = True
                frmODASMOperation.ProgressBar1.Value = 0
                frmODASMOperation.ProgressBar1.Max = (rsFindRecord!Duration / rsFindRecord2!NoOfMonths)
                frmODASMOperation.ProgressBar1.Min = 1
                
                While Times <> rsFindRecord!Duration / rsFindRecord2!NoOfMonths
                    Set UpDateMaintenance = New ADODB.Recordset
                    UpDateMaintenance.Open "SELECT * FROM ODASMMaintenance WHERE SiteNo = '" & rsFindRecord!SiteNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    
                    If UpDateMaintenance.EOF And UpDateMaintenance.BOF Then
                        MaintenanceNo = rsFindRecord!SiteNo & "- M" & "01"
                    ElseIf UpDateMaintenance.RecordCount <= 8 Then
                        MaintenanceNo = rsFindRecord!SiteNo & "- M0" & UpDateMaintenance.RecordCount + 1
                    ElseIf UpDateMaintenance.RecordCount >= 9 Then
                        MaintenanceNo = rsFindRecord!SiteNo & "- M" & UpDateMaintenance.RecordCount + 1
                    End If
                    
                        UpDateMaintenance.AddNew
                        UpDateMaintenance!MaintenanceNo = MaintenanceNo
                        UpDateMaintenance!SiteNo = rsFindRecord!SiteNo
                        UpDateMaintenance!MaintenanceDate = DateAdd("M", Period, rsFindRecord!MaintananceDueDate)
                        UpDateMaintenance!CurrentPeriod = CurrentPeriod
                        UpDateMaintenance!DatePrepared = Date
                        UpDateMaintenance!Maintained = "N"
                        UpDateMaintenance!PreparedBY = CurrentUserName
                        UpDateMaintenance.Update
                        Times = Times + 1: Period = Period + rsFindRecord2!NoOfMonths
                        frmODASMOperation.ProgressBar1.Value = frmODASMOperation.ProgressBar1.Value + 1
                        
                Wend
                frmODASMOperation.ProgressBar1.Visible = False
Continue:
                rsFindRecord!ComDate = Date
                If rsFindRecord!DurationMode = "D" Then
                rsFindRecord!ExDate = DateAdd("d", rsFindRecord!Duration, Date)
                ElseIf rsFindRecord!DurationMode = "M" Then
                rsFindRecord!ExDate = DateAdd("M", rsFindRecord!Duration, Date)
                ElseIf rsFindRecord!DurationMode = "Y" Then
                rsFindRecord!ExDate = DateAdd("yyyy", rsFindRecord!Duration, Date)
                End If
                rsFindRecord.Update
        
        Set rsFindRecord1 = New ADODB.Recordset
        rsFindRecord1.Open "SELECT * FROM ODASPPlotsite WHERE SiteNo = '" & rsFindRecord!SiteNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord1.EOF And rsFindRecord1.BOF Then Exit Sub
        rsFindRecord1!RateDueDate = rsFindRecord!ComDate
        rsFindRecord1!JCStartDate = Date
        rsFindRecord1!RateDue = rate
        rsFindRecord1!JCExpiryDate = rsFindRecord!ExDate
        rsFindRecord1.Update
            
        Set rsSchedule = New ADODB.Recordset
        rsSchedule.Open "Select * From ODASMSiteSchedule Where SiteNo = '" & rsFindRecord1!SiteNo & "' and JobBriefItemNo = '" & rsFindRecord1!JobBriefItemNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        frmODASMOperation.ProgressBar1.Visible = False
        frmODASMOperation.ProgressBar2.Visible = True
        frmODASMOperation.ProgressBar2.Min = 0
        frmODASMOperation.ProgressBar2.Max = rsSchedule.RecordCount
        frmODASMOperation.ProgressBar2.Value = 0
        
        rsSchedule.MoveFirst
        Do While Not rsSchedule.EOF
            If ScheduleDate <= rsFindRecord1!JCExpiryDate Then
                rsSchedule!Allocated = "Y"
            Else
                rsSchedule!Allocated = "N"
                rsSchedule!Reserved = "N"
            End If
                rsSchedule.Update
            rsSchedule.MoveNext
        frmODASMOperation.ProgressBar2.Value = frmODASMOperation.ProgressBar2.Value + 1
        Loop
        
       frmODASMOperation.ProgressBar2.Visible = False
       
    rsFindRecord.MoveNext
    Loop
    
    Set rsFindRecord = Nothing: Set rsFindRecord1 = Nothing
    
    Exit Sub
err:
ErrorMessage
End Sub
Public Sub invoiceJOBBRIEF()
On Error GoTo err
        
        Set rsCONTROL = New Recordset
        
        strCONTROL = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & CurrentRecord & "'"
        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
        
        If rsSAVE.EOF Or rsSAVE.BOF Then
                rsSAVE!Status = "INVOICE"
                rsSAVE!Invoiced = "Y"
                rsSAVE!DateINVOICED = Date
                rsSAVE!InvoicedBy = CurrentUserName
                rsSAVE.Update
                rsSAVE.Requery
        End If

rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub GenerateMaintenanceJobBriefNo()
On Error GoTo err
    Set rsFindRecord = New Recordset
    rsFindRecord.Open "SELECT * FROM ODASMMaintenance WHERE SiteNo = '" & frmODASMJobBrief.txtSiteNo.Text & "' and MaintenanceDate = '" & Format(frmODASMJobBrief.txtMaitenanceDate.Text, "MMMM dd,yyyy") & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    frmODASMJobBrief.txtJobBriefNo.Text = frmODASMJobBrief.txtJobBriefNo.Text & "-" & Right(rsFindRecord!MaintenanceNo, 3)
    frmODASMJobBrief.txtJobBriefItemNo.Text = frmODASMJobBrief.txtJobBriefItemNo.Text & "-" & Right(rsFindRecord!MaintenanceNo, 3)
    Set rsFindRecord = Nothing
Exit Sub
err:
ErrorMessage
End Sub
Public Function CheckDates() As Boolean
On Error GoTo err
    With frmODASMJobBrief
    
    Set rsFindRecord = New ADODB.Recordset: Set rsFindRecord1 = New ADODB.Recordset
    
    CheckDates = False
    
        If .optSupply.Value = True Then
            CheckDates = True: Exit Function
        Else
            
            
            If bBillBoard = True Then
                Set rsFindRecord2 = New ADODB.Recordset
                
                rsFindRecord2.Open "Select * From ODASPPlotSite Where MastNo = '" & .txtSiteNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsFindRecord2.EOF And rsFindRecord2.BOF Then Exit Function
                
                rsFindRecord2.MoveFirst
                
                Do While rsFindRecord2.EOF <> True
                
                    rsFindRecord1.Open "Select max(ScheduleDate)as LastDate,SiteNo From ODASMSiteSchedule Where SiteNo = '" & rsFindRecord2!SiteNo & "' Group By  SiteNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
                    rsFindRecord.Open "Select * From ODASMSiteSchedule Where SiteNo = '" & rsFindRecord2!SiteNo & "' and ScheduleDate = '" & Format(.txtExpectedDOC.Text, "MMMM dd,yyyy") & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                    If Trim(.cboDurationMode.Text) = "D" Then
                        ExDate = DateAdd("d", .txtPeriod.Text, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
                        If rsFindRecord1!LastDate < ExDate Then
                            MsgBox Message, vbCritical + vbOKOnly, "Out of Range JobBrief Dates"
                            Exit Function
                        Else
                            CheckDates = True
                        End If
                    ElseIf Trim(.cboDurationMode.Text) = "M" Then
                        ExDate = DateAdd("M", .txtPeriod.Text, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
                        If rsFindRecord1!LastDate < ExDate Then
                            MsgBox Message, vbCritical + vbOKOnly, "Out of Range JobBrief Dates"
                            Exit Function
                        Else
                            CheckDates = True
                        End If
                    ElseIf Trim(.cboDurationMode.Text) = "Y" Then
                        ExDate = DateAdd("yyyy", .txtPeriod.Text, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
                        If rsFindRecord1!LastDate < ExDate Then
                            MsgBox Message, vbCritical + vbOKOnly, "Out of Range JobBrief Dates"
                            Exit Function
                        Else
                            CheckDates = True
                        End If
                    End If
                    
                    If rsFindRecord!Reserved = "Y" Or rsFindRecord!Allocated = "Y" Then
                        If MsgBox("Site Chosen is Not free on this Date... Do you want System to propose Dates?", vbCritical + vbYesNo, "Site / Adverts Monitor") = vbYes Then
                             .cmdView.Caption = "View Dates"
                             CheckDates = False
                    
                    End If
                    End If
                rsFindRecord2.MoveNext
                Loop
            Else
                rsFindRecord1.Open "Select max(ScheduleDate)as LastDate,SiteNo From ODASMSiteSchedule Where SiteNo = '" & .txtSiteNo & "' Group By  SiteNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
                rsFindRecord.Open "Select * From ODASMSiteSchedule Where SiteNo = '" & .txtSiteNo.Text & "' and ScheduleDate = '" & Format(.txtExpectedDOC.Text, "MMMM dd,yyyy") & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If Trim(.cboDurationMode.Text) = "D" Then
                    ExDate = DateAdd("d", .txtPeriod.Text, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
                    If rsFindRecord1!LastDate < ExDate Then
                        MsgBox Message, vbCritical + vbOKOnly, "Out of Range JobBrief Dates"
                        Exit Function
                    Else
                        CheckDates = True
                    End If
                ElseIf Trim(.cboDurationMode.Text) = "M" Then
                    ExDate = DateAdd("M", .txtPeriod.Text, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
                    If rsFindRecord1!LastDate < ExDate Then
                        MsgBox Message, vbCritical + vbOKOnly, "Out of Range JobBrief Dates"
                        Exit Function
                    Else
                        CheckDates = True
                    End If
                ElseIf Trim(.cboDurationMode.Text) = "Y" Then
                    ExDate = DateAdd("yyyy", .txtPeriod.Text, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
                    If rsFindRecord1!LastDate < ExDate Then
                        MsgBox Message, vbCritical + vbOKOnly, "Out of Range JobBrief Dates"
                        Exit Function
                    Else
                        CheckDates = True
                    End If
                End If
                
                If rsFindRecord!Reserved = "Y" Or rsFindRecord!Allocated = "Y" Then
                    If MsgBox("Site Chosen is Not free on this Date... Do you want System to propose Dates?", vbCritical + vbYesNo, "Site / Adverts Monitor") = vbYes Then
                         .cmdView.Caption = "View Dates"
                         CheckDates = False
                
                End If
                End If
            End If
        End If
    End With
Exit Function
err:
ErrorMessage
End Function
Public Sub GenerateJobBriefNo()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASPLastNumbers WHERE AutoJobBriefNo = 'Y'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsCONTROL
                If .BOF Or .EOF = True Then Exit Sub
                        Screen.ActiveForm.txtJobBriefNo.Text = !JobBriefNo & ""
                
                Select Case Len(Trim(frmODASMJobBrief.txtJobBriefNo))
                        Case 1: frmODASMJobBrief.txtJobBriefNo.Text = Trim(!JobBriefPrefix) + "00000" + Trim(frmODASMJobBrief.txtJobBriefNo)
                        Case 2: frmODASMJobBrief.txtJobBriefNo.Text = Trim(!JobBriefPrefix) + "0000" + Trim(frmODASMJobBrief.txtJobBriefNo)
                        Case 3: frmODASMJobBrief.txtJobBriefNo.Text = Trim(!JobBriefPrefix) + "000" + Trim(frmODASMJobBrief.txtJobBriefNo)
                        Case 4: frmODASMJobBrief.txtJobBriefNo.Text = Trim(!JobBriefPrefix) + "00" + Trim(frmODASMJobBrief.txtJobBriefNo)
                        Case 5: frmODASMJobBrief.txtJobBriefNo.Text = Trim(!JobBriefPrefix) + "0" + Trim(frmODASMJobBrief.txtJobBriefNo)
                        Case 6: frmODASMJobBrief.txtJobBriefNo.Text = Trim(!JobBriefPrefix) + Trim(frmODASMJobBrief.txtJobBriefNo)
                End Select
                
                !JobBriefNo = !JobBriefNo + 1
                .Update
                .Requery

        End With
        
rsCONTROL.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub GenerateJobItemNo()
On Error GoTo err

        Set rsSAVE = New ADODB.Recordset
        
        strSQL = "SELECT * FROM ODASMJobBriefItems WHERE JobBriefNo = '" & frmODASMJobBrief.txtJobBriefNo.Text & "'"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsSAVE
                If .BOF Or .EOF = True Then
                        frmODASMJobBrief.txtJobBriefItemNo.Text = 1
                 Else
                        Screen.ActiveForm.txtJobBriefItemNo.Text = .RecordCount + 1
                End If
                        
                Select Case Len(Trim(frmODASMJobBrief.txtJobBriefItemNo))
                        Case 1: frmODASMJobBrief.txtJobBriefItemNo.Text = Trim(frmODASMJobBrief.txtJobBriefNo) + "-0" + Trim(frmODASMJobBrief.txtJobBriefItemNo)
                        Case 2: frmODASMJobBrief.txtJobBriefItemNo.Text = Trim(frmODASMJobBrief.txtJobBriefNo) + "-" + Trim(frmODASMJobBrief.txtJobBriefItemNo)
                End Select
                
'                !JobBriefItemNo = frmODASMJobBrief.txtJobBriefNo.Text
'                .Update
'                .Requery

        End With
        
rsSAVE.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub saveJobBrief()
On Error GoTo err
       With frmODASMJobBrief
            
                strSQL = "select * from ODASMJobBrief Where JobBriefNo = '" & frmODASMJobBrief.txtJobBriefNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobBriefNo = Trim(.txtJobBriefNo.Text)
                        rsSAVE!CreatedBy = CurrentUserName
                        rsSAVE!DateCreated = Date
                        rsSAVE!Created = "Y"
                        rsSAVE!Approved = "N"
                        rsSAVE!Authorized = "N"
                        rsSAVE!Closed = "N"
                        rsSAVE!items = 0
                        rsSAVE!opened = "N"
                        rsSAVE!Invoiced = "N"
                        rsSAVE!Cancelled = "N"
                        rsSAVE!Started = "N"
                        rsSAVE!Status = "JOBBRIEF-CREATED"
                        rsSAVE!creditAuthorized = "N"
                End If
                
                rsSAVE!BusinessSource = .cboSource.Text
                rsSAVE!descriptionOfOrder = .txtDescriptionOfOrder.Text
                rsSAVE!ProductCode = .txtProductCode.Text
                rsSAVE!ReceivedBy = .txtReceivedBy.Text
                rsSAVE!BriefBy = .txtBriefBy.Text
                If .optSupply.Value = True Then
                    rsSAVE!ContractPeriod = 0
                Else
                    rsSAVE!ContractPeriod = .txtPeriod.Text
                End If
                rsSAVE!PaymentMode = .cboPaymentMode.Text
                rsSAVE!PaymentMethod = .cboPaymentMethod.Text
                rsSAVE!JobBriefDate = .txtJobBriefDate.Text
                rsSAVE!expectedDOC = .txtExpectedDOC.Text
                rsSAVE!QuotationNumber = .txtQuotationNo.Text
                rsSAVE!AccountNo = .txtAccountNo.Text
                rsSAVE!deadlineDate = .txtdeadlineDate.Text
                rsSAVE!AccPeriod = CurrentPeriod
                rsSAVE!AdditionalInformation = .txtAdditionalInformation.Text
                
                
                rsSAVE.Update
                rsSAVE.Requery
                
                Set rsNewRecord = Nothing
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub saveMaintenanceJobBrief()
On Error GoTo err
       With frmODASMJobBrief
            
                strSQL = "select * from ODASMJobBrief Where JobBriefNo = '" & frmODASMJobBrief.txtJobBriefNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobBriefNo = Trim(.txtJobBriefNo.Text)
                        rsSAVE!CreatedBy = CurrentUserName
                        rsSAVE!DateCreated = Date
                        rsSAVE!Created = "Y"
                        rsSAVE!Approved = "N"
                        rsSAVE!Authorized = "N"
                        rsSAVE!Closed = "N"
                        rsSAVE!items = 0
                        rsSAVE!opened = "N"
                        rsSAVE!Invoiced = "N"
                        rsSAVE!Cancelled = "N"
                        rsSAVE!Started = "N"
                        rsSAVE!Status = "JOBBRIEF-CREATED"
                        rsSAVE!creditAuthorized = "N"
                End If
                
                
                rsSAVE!descriptionOfOrder = .txtDescriptionOfOrder.Text & ""
                rsSAVE!ProductCode = .txtProductCode.Text & ""
                rsSAVE!ReceivedBy = .txtReceivedBy.Text & ""
                rsSAVE!JobBriefDate = Date
                rsSAVE!expectedDOC = .txtMaitenanceDate.Text
                rsSAVE!AccountNo = .txtAccountNo.Text & ""
                rsSAVE!deadlineDate = .txtMaitenanceDate.Text
                rsSAVE!AccPeriod = CurrentPeriod
                rsSAVE!AdditionalInformation = .txtAdditionalInformation.Text & ""
                
                
                rsSAVE.Update
                rsSAVE.Requery
                
                
                Set rsNewRecord = Nothing
        
        End With
Exit Sub

err:
    ErrorMessage

End Sub

Public Sub updateQuotationItems()
On Error GoTo err
    
       With frmODASMJobBrief
                strSQL = "select * from ODASMQuotationItems Where QuotationItemNo = '" & frmODASMJobBrief.txtQuotationItemNo.Text & "' ;"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                        rsSAVE!Status = "BRIEFED"
                        rsSAVE.Update
                        rsSAVE.Requery
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage

End Sub
Public Sub saveMaintenanceJobBriefItem()
On Error GoTo err
    
       With frmODASMJobBrief

                strSQL = "select * from ODASMJobBriefItems Where JobBriefNo = '" & frmODASMJobBrief.txtJobBriefNo & "' and JobBriefItemNo = '" & frmODASMJobBrief.txtJobBriefItemNo.Text & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobBriefNo = .txtJobBriefNo.Text
                        rsSAVE!JobBriefItemNo = .txtJobBriefItemNo.Text
                        rsSAVE!CreatedBy = CurrentUserName
                        rsSAVE!DateCreated = Date
                End If
                        rsSAVE!MediaCode = .txtMediaCode.Text
                        rsSAVE!MediaSize = .txtMediaSize.Text
                        rsSAVE!SiteNo = .txtSiteNo.Text
                        rsSAVE!Town = .txtTown.Text
                        rsSAVE!PhysicalLocation = .txtPhysicalLocation.Text
                        
                        rsSAVE.Update
                        
                        FindSQL = "select * from ODASMMaintenance Where MaintenanceDate = '" & Format(.txtMaitenanceDate.Text, "MMMM dd,yyyy") & "' and SiteNo = '" & .txtSiteNo.Text & "';"
                        Set rsFIND = New ADODB.Recordset
                        rsFIND.Open FindSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        If rsFIND.EOF Or rsFIND.BOF Then Exit Sub
                                rsFIND!JobBriefCreated = "Y"
                                rsFIND.Update
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage

End Sub
Public Sub saveSiteSchedule()
On Error GoTo err
    With frmODASMJobBrief
    
        Set rsFindRecord1 = New ADODB.Recordset
        rsFindRecord1.Open "SELECT * FROM ODASMJobBriefItems WHERE JobBriefItemNo = '" & .txtJobBriefItemNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        Dim EDate As Date
        
        If rsFindRecord1!DurationMode = "D" Then
            EDate = DateAdd("d", rsFindRecord1!Duration, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
        ElseIf rsFindRecord1!DurationMode = "M" Then
            EDate = DateAdd("M", rsFindRecord1!Duration, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
        ElseIf rsFindRecord1!DurationMode = "Y" Then
            EDate = DateAdd("yyyy", rsFindRecord1!Duration, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
        End If

        Set rsFindRecord = New ADODB.Recordset
        
        If bBillBoard = True And .optRenting.Value = True Then
            rsFindRecord.Open "SELECT * FROM ODASPPlotSite WHERE MastNo = '" & .txtSiteNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        Else
            rsFindRecord.Open "SELECT * FROM ODASPPlotSite WHERE SiteNo = '" & .txtSiteNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        End If
        
        If rsFindRecord.RecordCount = 0 Then Exit Sub
        
        Set rsSITE = New ADODB.Recordset
        
        rsFindRecord.MoveFirst
        
        Do While rsFindRecord.EOF <> True
            rsSITE.Open "SELECT * FROM ODASMSiteSchedule WHERE SiteNo ='" & rsFindRecord!SiteNo & "' and (ScheduleDate between '" & Format(rsFindRecord1!CommencementDate, "MMMM dd,yyyy") & "' And '" & Format(EDate, "MMMM dd,yyyy") & "') ;", cnCOMMON, adOpenKeyset, adLockOptimistic
            .ProgressBar1.Visible = True: .ProgressBar1.Value = 0
            .ProgressBar1.Max = rsSITE.RecordCount
            .ProgressBar1.Min = 0
            rsSITE.MoveFirst
            
            Do While Not rsSITE.EOF
                rsSITE!JobBriefItemNo = .txtJobBriefItemNo.Text
                rsSITE!Reserved = "Y"
                rsSITE.Update
            
            rsSITE.MoveNext
            .ProgressBar1.Value = .ProgressBar1.Value + 1
            Loop
            .ProgressBar1.Visible = False
        rsFindRecord.MoveNext
        Loop
        .ProgressBar1.Visible = False
    End With

Exit Sub
err:
ErrorMessage
End Sub
Public Sub ComputeCouncilRatesPayable()
On Error GoTo err
    With frmODASMJobBrief
    
        Dim LDate, UDate As Variant
        
        Set rsFindRecord1 = New ADODB.Recordset
        Dim rate As Double
        rsFindRecord1.Open "SELECT * FROM ODASMJobBriefItems WHERE JobBriefItemNo = '" & .txtJobBriefItemNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
       
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPPlotSite WHERE SiteNo = '" & .txtSiteNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.RecordCount = 0 Then Exit Sub
        
        If .cboTown.Text = "NRB" Then
        
            Set rsNairobiRatesPaid = New ADODB.Recordset
            rsNairobiRatesPaid.Open "Select Max(ODASMSiteSchedule.ScheduleDate)as EDate,ODASMSiteSchedule.* From ODASMSiteSchedule Where SiteNo = '" & .txtSiteNo.Text & "' and JobBriefItemNo='" & .txtJobBriefItemNo.Text & "' and LicenceNo is not NULL ;", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Set rsNairobiRates = New ADODB.Recordset
            rsNairobiRates.Open "Select * From ODASMSiteSchedule Where SiteNo = '" & .txtSiteNo.Text & "' and JobBriefItemNo='" & .txtJobBriefItemNo.Text & "' and LicenceNo is NULL ;", cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsNairobiRates.RecordCount = 0 Then
                    rate = 0
                Else
                    rsNairobiRates.MoveFirst
                        LDate = rsNairobiRates!ScheduleDate
                    rsNairobiRates.MoveLast
                        UDate = rsNairobiRates!ScheduleDate
                    
                    If rsFindRecord1!DurationMode = "M" And DateDiff("M", LDate, UDate) <= 3 Then
                        rate = CDbl(rsFindRecord!Rates) / 4
                    ElseIf rsFindRecord1!DurationMode = "M" And DateDiff("M", LDate, UDate) > 3 And rsFindRecord1!Duration <= 6 Then
                        rate = CDbl(rsFindRecord!Rates) / 2
                    ElseIf rsFindRecord1!DurationMode = "M" And DateDiff("M", LDate, UDate) > 6 And DateDiff("M", LDate, UDate) <= 12 Then
                        rate = CDbl(rsFindRecord!Rates)
                    ElseIf rsFindRecord1!DurationMode = "M" And DateDiff("M", LDate, UDate) > 12 Then
                        rate = CDbl(rsFindRecord!Rates) / 12 * DateDiff("M", LDate, UDate)
                    End If
                    
                    If rsFindRecord1!DurationMode = "Y" Then
                        rate = CDbl(rsFindRecord!Rates)
                    End If
                    
                    If rsFindRecord1!DurationMode = "D" And DateDiff("d", LDate, UDate) <= 90 Then
                        rate = CDbl(rsFindRecord!Rates) / 4
                    ElseIf rsFindRecord1!DurationMode = "D" And (DateDiff("d", LDate, UDate) > 90 And DateDiff("d", LDate, UDate) <= 180) Then
                        rate = CDbl(rsFindRecord!Rates) / 2
                    End If
                 End If
        Else
            Set rsFindRatePaid = New ADODB.Recordset
            rsFindRatePaid.Open "Select Max(ODASMSiteSchedule.ScheduleDate)as EDate,ODASMSiteSchedule.* From ODASMSiteSchedule Where SiteNo = '" & .txtSiteNo.Text & "' and JobBriefItemNo='" & .txtJobBriefItemNo.Text & "' and LicenceNo is not NULL ;", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsFindRatePaid.RecordCount <> 0 Then
            
                Set rsFindRate = New ADODB.Recordset
                rsFindRate.Open "Select * From ODASMSiteSchedule Where SiteNo = '" & .txtSiteNo.Text & "' and JobBriefItemNo='" & .txtJobBriefItemNo.Text & "' and LicenceNo is not NULL ;", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsFindRate.RecordCount = 0 Then
                    rate = 0
                    
                    Else
                        rsFindRate.MoveFirst
                            LDate = rsFindRate!ScheduleDate
                        rsFindRate.MoveLast
                            UDate = rsFindRate!ScheduleDate
                            
                        If Year(LDate) = Year(EDate) Then
                            rate = (Year(UDate) - Year(LDate)) * rsFindRecord!Rates
                        ElseIf Year(LDate) > Year(EDate) Then
                            rate = ((Year(UDate) - Year(LDate)) + 1) * rsFindRecord!Rates
                        End If
                    
                End If
            Else
                rate = ((Year(UDate) - Year(LDate)) + 1) * rsFindRecord!Rates
            End If
            
        End If
        
        Set rsSITE = New ADODB.Recordset
        rsSITE.Open "SELECT * FROM ODASMCouncilRatesPayable WHERE SiteNo ='" & rsFindRecord!SiteNo & "' and JobBriefItemNo= '" & .txtJobBriefItemNo.Text & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsSITE.RecordCount = 0 Then
            rsSITE.AddNew
            rsSITE!JobBriefItemNo = .txtJobBriefItemNo.Text
            rsSITE!SiteNo = .txtSiteNo.Text
        End If
            rsSITE!TownCode = .cboTown.Text
            rsSITE!Duration = rsFindRecord1!Duration & " " & rsFindRecord1!DurationMode
            rsSITE!RatePayable = rate
            rsSITE.Update
        
    End With

Exit Sub
err:
ErrorMessage
End Sub

Public Sub saveJobBriefItems()
On Error GoTo err
    
       With frmODASMJobBrief
                Dim strIlluminated, strBordered, Supply, Contract As String
            
                If .chkBorder.Value = True Then strBordered = "Y" Else strBordered = "N"
                If .chkIlluminate.Value = True Then strIllumintaed = "Y" Else strIlluminated = "N"
                If .optContract.Value = True Then Contract = "Y" Else Contract = "N"
                If .optSupply.Value = True Then Supply = "Y" Else Supply = "N"
                
                strSQL = "select * from ODASMJobBriefItems Where JobBriefNo = '" & frmODASMJobBrief.txtJobBriefNo & "' and JobBriefItemNo = '" & frmODASMJobBrief.txtJobBriefItemNo.Text & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobBriefNo = .txtJobBriefNo.Text
                        rsSAVE!JobBriefItemNo = .txtJobBriefItemNo.Text
                        rsSAVE!CreatedBy = CurrentUserName
                        rsSAVE!DateCreated = Date
                End If
                        rsSAVE!MediaCode = .txtMediaCode.Text
                        rsSAVE!MediaSize = .txtMediaSize.Text
                        If .txtQuotationNo.Text <> Empty Then
                            Set rsFindRecord = New ADODB.Recordset
                            rsFindRecord.Open "SELECT * FROM ODASMQuotationItems WHERE QuotationItemNo = '" & .txtQuotationItemNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                               
                            rsSAVE!ItemPrice = rsFindRecord!ItemPrice
                            rsSAVE!PriceExclusive = rsFindRecord!PriceExclusive
                            rsSAVE!PriceInclusive = rsFindRecord!PriceInclusive
                            rsSAVE!VatAmount = rsFindRecord!VatAmount
                        Else
                            
                            rsSAVE!ItemPrice = .txtprice.Text
                            If bBillBoard = True Then
                                If .optRenting.Value = True Then
                                    rsSAVE!PriceExclusive = .txtAmount.Text
                                    rsSAVE!PriceInclusive = .txtTotalCost.Text
                                    rsSAVE!VatAmount = .txtTax.Text
                                End If
                            Else
                                rsSAVE!PriceExclusive = .txtAmount.Text
                                rsSAVE!PriceInclusive = .txtAmountInc.Text
                                rsSAVE!VatAmount = .txtTax.Text
                            End If
                        End If
                        rsSAVE!SiteNo = .txtSiteNo.Text
                        rsSAVE!Town = .txtTown.Text & ""
                        rsSAVE!illuminated = strIlluminated
                        rsSAVE!SidingCode = .cboSidingCode.Text
                        rsSAVE!ColorCode = .cboColorCode.Text
                        If .optSupply.Value = True Then
                            rsSAVE!Duration = 0
                        Else
                            rsSAVE!Duration = .txtPeriod.Text
                        End If
                        rsSAVE!CommencementDate = Format(.txtExpectedDOC.Text, "MMMM dd,yyyy")
                        rsSAVE!DurationMode = Trim(.cboDurationMode.Text) & ""
                        rsSAVE!ItemQuantity = Val(.txtItemQuantity.Text)
                        rsSAVE!PhysicalLocation = .txtPhysicalLocation.Text
                        rsSAVE!Contract = Contract
                        rsSAVE!Supply = Supply
                        rsSAVE!Bordered = strBordered
                        If .txtSiteNo.Text = "N/A" Then
                            rsSAVE!RequireSite = "N"
                        Else
                            If .optSupply.Value = True Then
                                rsSAVE!RequireSite = "N"
                                
                            Else
                                 rsSAVE!Duration = .txtPeriod.Text & ""
                                 rsSAVE!MaintananceInterval = .cboMaintanancePeriod.Text & ""
                            End If
                            
                        End If
                        If .chkAcquireSite.Value = 1 Then
                                rsSAVE!AcquireSite = "Y"
                        Else: rsSAVE!AcquireSite = "N"
                        End If
                        
                        If .optContract.Value = True Then
                            If Trim(.cboDurationMode.Text) = "D" Then
                                rsSAVE!ExpiryDate = DateAdd("d", .txtPeriod.Text, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
                            ElseIf Trim(.cboDurationMode.Text) = "M" Then
                                rsSAVE!ExpiryDate = DateAdd("M", .txtPeriod.Text, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
                            ElseIf Trim(.cboDurationMode.Text) = "Y" Then
                                rsSAVE!ExpiryDate = DateAdd("yyyy", .txtPeriod.Text, Format(.txtExpectedDOC.Text, "MMMM dd,yyyy"))
                            End If
                        End If
                        
                        rsSAVE.Update
                        
                        If .optSupply.Value = True Then
                            'ComputeJobBriefPrices
                        End If
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub saveJobBriefATTACHMENT()
On Error GoTo err
    
       With frmODASMJobBrief

                strSQL = "select * from ODASMJobBriefAttachment Where JobBriefNo = '" & frmODASMJobBrief.txtJobBriefNo & "' and Attachmentcode = '" & .txtAttachmentCode.Text & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobBriefNo = .txtJobBriefNo.Text
                        rsSAVE!AttachmentCode = .txtAttachmentCode.Text
                        rsSAVE!PreparedBY = CurrentUserName
                        rsSAVE!DatePrepared = Date
                End If
                
                rsSAVE.Update
                rsSAVE.Requery
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub updateSITE()
On Error GoTo err
    
        ' procedure is used to update the site on the Bill Board. If the Bill Board has two sides
        ' then both are updared to show that they are available
        
        
      With frmODASMJobBrief
                
                If bBillBoard = True And .optRenting.Value = True Then
                    
                    strCONTROL = "select * from ODASPPlotSite, ODASPPlotMast Where ODASPPlotSite.MastNo = '" & .txtSiteNo & "' and ODASPPlotSite.status = 'SITE-AVAILABLE' and ODASPPlotSite.MastNo = ODASPPlotMast.MastNo;"
                    Set rsCONTROL = New ADODB.Recordset
                    rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Else
                    strCONTROL = "select * from ODASPPlotSite, ODASPPlotMast Where ODASPPlotSite.SiteNo = '" & .txtSiteNo & "' and ODASPPlotSite.status = 'SITE-AVAILABLE' and ODASPPlotSite.MastNo = ODASPPlotMast.MastNo;"
                    Set rsCONTROL = New ADODB.Recordset
                    rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
               
               End If
                    
                    Do While Not rsCONTROL.EOF
                    
                        Set rsSAVE = New ADODB.Recordset
                        strSQL = "select * from ODASPPlotSite Where ODASPPlotSite.MastNo = '" & rsCONTROL!MastNo & "' and ODASPPlotSite.status = 'SITE-AVAILABLE';"
                        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        If rsSAVE.EOF And rsSAVE.BOF Then
                        Else
'                                rsSAVE!Status = "SITE-RESERVED"
                                rsSAVE!JobBriefItemNo = .txtJobBriefItemNo
                                rsSAVE!JobBriefNo = .txtJobBriefNo.Text
                                rsSAVE!ReservationDate = .txtJobBriefDate.Text
                                rsSAVE!MediaSize = .txtMediaSize.Text
                                rsSAVE!MediaCode = .txtMediaCode.Text
                                rsSAVE.Update
                        End If
                        rsCONTROL.MoveNext
                Loop
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateMAST()
On Error GoTo err
    
        ' procedure is used to update the mast which the Bill Board is installed on. If the Bill Board has two sides
        ' then both are updared to show that they are available
        
        
    With frmODASMJobBrief
                
            Set rsSAVE = New ADODB.Recordset
            strSQL = "select * from ODASPPlotMast Where ODASPPlotMast.MastNo = '" & .txtSiteNo.Text & "' and ODASPPlotMAST.status = 'SITE-AVAILABLE';"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
            If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
            
'                rsSAVE!Status = "MAST-RESERVED"
                rsSAVE!JobBriefItemNo = .txtJobBriefItemNo
'                rsave!JobBriefNo = .txtJobBriefItemNo.Text
                rsSAVE!ReservationDate = .txtJobBriefDate.Text
                rsSAVE.Update
                rsSAVE.Requery
                
            Set rsSAVE = Nothing
            strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableRECORD()
On Error GoTo err
        With frmODASMJobBrief
            .Frame10.Enabled = True: .optContract.Enabled = True: .optSupply.Enabled = True
            .txtAdditionalInformation.Locked = False
            .txtCompanyName.Locked = True
            .txtBriefBy.Locked = True
            .UpDown1.Enabled = True
            .txtProductCode.Locked = True
            .UpDownQuantity.Enabled = True
            .txtReceivedBy.Locked = True
            .txtDescriptionOfOrder.Locked = False
            .txtExpectedDOC.Locked = False
            .txtItemQuantity.Locked = False
            .txtMediaCode.Locked = False
            .txtMediaSize.Locked = False
            .cboColorCode.Locked = False
            .cboSidingCode.Locked = False
            .chkAcquireSite.Enabled = True
            .chkBorder.Enabled = True
            .chkIlluminate.Enabled = True
            .txtBriefBy.Locked = False
            .txtProductCode.Locked = False
            .txtReceivedBy.Locked = False
            .txtAttachmentCode.Locked = False
            .dtpDeadlineDate.Enabled = True
            .DTPickerExpectedDOC.Enabled = True
            .txtSiteLocation.Locked = False
            .txtPhysicalLocation.Locked = False
            .cmdView.Enabled = True
            .cboPaymentMethod.Locked = False
            .cboPaymentMode.Enabled = True
            .cboPaymentMode.Locked = False
            .cboDurationMode.Enabled = True
            .cboDurationMode.Locked = False
            .cboTown.Enabled = True
            .cboSource.Locked = False
            .cboTown.Locked = False: .txtprice.Text = ""
            .cboMaintanancePeriod.Locked = False
            .cboMaintanancePeriod.Enabled = True
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Function ValidRecord() As Boolean
With frmODASMJobBrief
    
    If .txtContactName.Text = Empty Then
        strMessage = "Name Of Contact Person Required...!"
        .txtContactName.SetFocus
    ElseIf .optContract.Value = True And .txtPeriod.Text = Empty Then
        strMessage = "Required Period of Contract..."
        .txtPeriod.SetFocus: .UpDown1.Enabled = True
    ElseIf .optContract.Value = True And .txtTown.Text = Empty Then
        strMessage = "Town of the Advertisement required...!"
        .txtTown.SetFocus
    ElseIf .cboSource.Text = Empty Then
        strMessage = "Source of the Business required...!"
        .cboSource.SetFocus
    ElseIf .cboSidingCode.Text = Empty Then
        strMessage = "Advertisement siding required...!"
        .cboSidingCode.SetFocus
    ElseIf .cboPaymentMethod.Text = Empty Then
        strMessage = "Payment Method has to be specified...!"
        .cboPaymentMethod.SetFocus
    ElseIf .optContract.Value = True And Trim(.cboDurationMode.Text) = Empty Then
        strMessage = "Duration Mode has to be specified...!"
        .cboDurationMode.SetFocus
    ElseIf .optContract.Value = True And .cboMaintanancePeriod.Text = Empty Then
        strMessage = "Maintanance Period has to be specified...!"
        .cboMaintanancePeriod.SetFocus
    ElseIf .cboPaymentMode.Text = Empty Then
        strMessage = "Payment Method has to be specified...!"
        .cboPaymentMode.SetFocus
    ElseIf .txtBriefBy.Text = Empty Then
        strMessage = "Briefed by cannot be Left Blank"
        .txtBriefBy.SetFocus
    ElseIf .txtProductCode.Text = Empty Then
        strMessage = "The Product Code Cannot be Blank"
        .txtProductCode.SetFocus
    ElseIf .txtReceivedBy.Text = Empty Then
        strMessage = "Received By Cannot be Blank"
        .txtReceivedBy.SetFocus
    ElseIf .txtJobBriefDate.Text = Empty Then
        Strmsssage = "The Job Brief Date is Blank"
        .txtJobBriefDate.SetFocus
    ElseIf .txtAccountNo.Text = Empty Then
        strMessage = "The Customer Has to be Selected"
        .txtAccountNo.SetFocus
    ElseIf .txtDescriptionOfOrder.Text = Empty Then
        strMessage = "The Description of Order is Required to Identify the Order"
        .txtDescriptionOfOrder.SetFocus
    ElseIf .txtdeadlineDate.Text = Empty Then
        strMessage = "The Deadline Date is Required"
        .txtdeadlineDate.SetFocus
    ElseIf .txtAttachmentCode.Text = Empty Then
        strMessage = "The Attachment Code is Necessary"
        .txtAttachmentCode.SetFocus
    ElseIf .txtExpectedDOC.Text = Empty Then
        strMessage = "The Expected Commencement Date is Necessary"
        .txtExpectedDOC.SetFocus
    ElseIf DateDiff("D", .txtExpectedDOC, .txtdeadlineDate) < 0 Then
        strMessage = "The expected DOC cannot be after the deadline Date"
        .txtdeadlineDate.SetFocus
    ElseIf .txtMediaCode.Text = Empty Then
        strMessage = "The Media Code cannot be Left Blank"
        .txtMediaCode.SetFocus
    ElseIf .txtMediaSize.Text = Empty Then
        strMessage = "The Media Size cannot be left Blank "
        .txtMediaSize.SetFocus
    ElseIf .txtItemQuantity.Text = Empty Then
        strMessage = "The Item Quantity is Required"
        .txtItemQuantity.SetFocus
    ElseIf .txtItemQuantity.Text <> Empty And .txtItemQuantity.Text <= 0 Then
        strMessage = "The Item Quantity is Required"
        .txtItemQuantity.SetFocus
    Else
        ValidRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
End Function


Public Sub clearRECORD()
On Error GoTo err
        
        With frmODASMJobBrief
                .txtMediaSize.Text = ""
                .chkIlluminate.Value = 0
                .txtprice.Text = ""
                .cboSidingCode.Text = Empty
                .chkBorder.Value = 0
                .txtItemQuantity.Text = 1
                .txtSiteName.Text = Empty
                .txtPhysicalLocation.Text = Empty
                .txtAttachmentCode.Text = Empty
                .txtJobBriefItemNo.Text = Empty
                .txtMediaSize.Text = Empty
                .txtSiteNo.Text = Empty
                .ListView3.ListItems.Clear
'                RemoveCurrentList3Item
'                RemoveCurrentList2Item
            
        End With

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub loadCONTACT()
On Error GoTo err
        
        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT * FROM ODASPAccount WHERE AccountNo = '" & Screen.ActiveForm.txtAccountNo.Text & "' ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
                    If .BOF And .EOF Then Exit Sub
                    Screen.ActiveForm.txtCompanyName.Text = !CompanyName
                    Screen.ActiveForm.txtContactName.Text = !ContactPerson
        End With

rsCONTROL.Close
strSQL = ""

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub loadDEFAULTPAYMENTMETHOD()
On Error GoTo err
        
        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT * FROM ODASPPaymentMethod WHERE AccountNoRequired = 'Y' ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
        frmODASMJobBrief.cboPaymentMethod.Text = rsCONTROL!PaymentMethod
                
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadDEFAULTS()
On Error GoTo err
    
    With frmODASMJobBrief
        .txtJobBriefDate.Text = Date
        .txtdeadlineDate.Text = Date
        .txtExpectedDOC.Text = Date
        .dtpBriefDate.Value = Date
        .cboPaymentMode.Text = "M"
        .dtpDeadlineDate.Value = Date
        .DTPickerExpectedDOC.Value = Date
        .txtItemQuantity.Text = 1
        .txtReceivedBy.Text = CurrentUserName
        '/Obtain The Default Payment Payment Method
        
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadCATEGORY()
On Error GoTo err
        
        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT * FROM AdvertCategories WHERE MediaCode = '" & Screen.ActiveForm.txtMediaCode.Text & "' ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
                    If .BOF And .EOF Then Exit Sub
                    Screen.ActiveForm.txtCategoryName.Text = !CategoryName
        End With

rsCONTROL.Close
strSQL = ""

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub calculatePRICE()
On Error GoTo err
        
        With frmODASMJobBrief
                
                .txtItemQuantity.Text = 1
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calculateTOTALQUANTITY()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT sum(ItemQuantity) as GrandQuantity FROM ODASMJobBriefItems WHERE JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then
'                frmODASMJobBrief.txtTotalQuantity.Text = 0
        ElseIf IsNull(rsCONTROL!GrandQuantity) = True Then
'                frmODASMJobBrief.txtTotalQuantity.Text = 0
'        Else: frmODASMJobBrief.txtTotalQuantity.Text = rsCONTROL!GrandQuantity
        End If


Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calculateTOTALPRICE()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT sum(NetItemPrice) as GrandPrice FROM ODASMJobBriefItems WHERE JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then
'                frmODASMJobBrief.txtTotalPrice.Text = 0
        ElseIf IsNull(rsCONTROL!GrandPrice) = True Then
'                frmODASMJobBrief.txtTotalPrice.Text = 0
'        Else: frmODASMJobBrief.txtTotalPrice.Text = rsCONTROL!GrandPrice
        End If


Exit Sub

err:
    ErrorMessage
End Sub


Public Sub loadSiteName()
On Error GoTo err
    With frmODASMJobBrief
        
        Set rsCONTROL = New ADODB.Recordset
        
        If rsMEDIA!RequireSite = "Y" Then
                strSQL = "SELECT * FROM ODASPPlotSite WHERE SiteNo = '" & .txtSiteNo.Text & "' ;"
        ElseIf rsMEDIA!RequirePlot = "Y" Then
             strSQL = "SELECT * FROM ODASPPlot, ODASPPlotSite WHERE ODASPPlotSite.SiteNo = '" & .txtSiteNo.Text & "' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ;"
        End If
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCONTROL.BOF And rsCONTROL.EOF Then Exit Sub
        
            If rsMEDIA!RequireSite = "Y" Then
                    .txtSiteName.Text = rsCONTROL!SiteDetails
            ElseIf rsMEDIA!RequirePlot = "Y" Then
                    .txtSiteName.Text = rsCONTROL!PlotName
            End If
'        End If
    End With

rsCONTROL.Close
strSQL = ""
finish:
Exit Sub
err:
    ErrorMessage
End Sub

