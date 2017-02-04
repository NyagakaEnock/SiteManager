VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASCouncilRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public rsJOBBRIEF, rsCOUNCIL, rsDUEDATE, rsSITE As ADODB.Recordset, strJOBBRIEF, StrCOUNCIL, strDUEDATE, strSITE As String
Public dtExpiryDate, dtStartDate As Date
Public intDURATION As Integer
Public strPAYNMENTMODE As String
Public Sub loadRECORD()
On Error GoTo err
        With frmODASMCouncilRates
            
                Set rsCONTROL = New ADODB.Recordset
                strCONTROL = "SELECT P.*,(PS.MediaSize) as Media,PS.*, PM.*,(PM.CommencementDate)as ComDate,(PM.ExpiryDate) as ExDate FROM ODASPPlot P,ODASPPlotmast PM, ODASPPlotSite PS WHERE PS.SiteNo = '" & .txtSiteNo & "' and PS.MastNo = PM.MastNo and PS.PlotNo = P.PlotNo;"
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic

                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                .txtPlotNo.Text = rsCONTROL!PLotNo
                .txtSiteDetails.Text = rsCONTROL!SiteDetails & ""
                .txtMediaCode.Text = rsCONTROL!MediaCode & ""
                .txtMediaSize.Text = rsCONTROL!Media & ""
                .txtCommencementDate.Text = rsCONTROL!ComDate & ""
                .txtExpiryDate.Text = rsCONTROL!ExDate & ""
                .txtAccountNo.Text = rsCONTROL!CouncilAccountNo & ""
                        
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadSTARTDATE()
On Error GoTo err
        With frmODASMCouncilRates
            
                Set rsCONTROL = New ADODB.Recordset
                strCONTROL = "SELECT * FROM ODASMJobBriefitems WHERE JobBriefItemNo = '" & .txtJobBriefItemNo & "' ;"
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic

                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If NewRecord = True Then
                        .txtCurrentYear.Text = CurrentYear
                        .txtStatus.Text = "REQUISITIONED"
                        
                        '/ obtain the Payment mode based on the Council Selected
                        
                        Set rsSAVE = New ADODB.Recordset
                        strSQL = "SELECT * FROM ODASPCouncil WHERE CouncilCode = '" & .txtTownCode.Text & "' ;"
                        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                        '/ obtain the first and the Last Day of the Year
                        
                        Dim rsYEAR As ADODB.Recordset, stryear As String
                        Set rsYEAR = New ADODB.Recordset
                        stryear = "SELECT * FROM ODASPYear WHERE CurrentYear = '" & .txtCurrentYear.Text & "' ;"
                        rsYEAR.Open stryear, cnCOMMON, adOpenKeyset, adLockOptimistic


                        If rsSAVE.EOF Or rsSAVE.BOF Then
                                MsgBox "Council Not Set up Yet"
                                Exit Sub
                        End If
                        
                        
                        
                        '/ Check whether the Yearly Council Rate Have been Set
                        
                        If rsSAVE!UseCalendarYear = "Y" Then
                                .cboPaymentMode.Text = "Y"
                                .txtDuration.Text = 12
                                .cboPaymentMode.Enabled = False
                                .cboPaymentMode.BackColor = &HFFFFC0
                                .txtRateStartDate.Text = CDate(rsYEAR!StartDate)
                                .txtRateExpiryDate.Text = CDate(rsYEAR!EndDate)

                        Else
                                .cboPaymentMode.Text = "Q"
                                .txtDuration.Text = 3
                                .cboPaymentMode.BackColor = &HFFC0C0
                                .cboPaymentMode.Enabled = True
                                .txtRateStartDate.Text = CDate(rsCONTROL!CommencementDate)
                                .txtRateStartDate.Text = DateAdd("D", -1 * (CDbl(Day(.txtRateStartDate)) - 1), .txtRateStartDate)
                                .txtRateExpiryDate.Text = DateAdd("M", CDbl(.txtDuration.Text), CDate(.txtRateStartDate))
                                
                                obtainLASTDAY
                                
                                If CDate(.txtRateExpiryDate) > CDate(dtLastDateOfYear) Then
                                        .txtRateExpiryDate.Text = CDate(dtLastDateOfYear)
                                End If
                                
                                .txtRateExpiryDate.Text = DateAdd("D", -1, CDate(.txtRateExpiryDate))

                        End If
                        
                End If
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadSTARTDATE_batch()
On Error GoTo err
            
            obtainLASTDAY
            
            '/ Open the all the sites
            
            Set rsCONTROL = New ADODB.Recordset
            strCONTROL = "SELECT * FROM ODASPPlotSite WHERE Active = 'Y' ;"
            rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                      
            frmODASMEOY.ProgressBar1.Visible = True
            frmODASMEOY.ProgressBar1.Value = 0
            frmODASMEOY.ProgressBar1.Max = rsCONTROL.RecordCount
                      
            Do While Not rsCONTROL.EOF
            
                        '/ Open the Rates Schedule File
                        
                        Set rsSAVE = New ADODB.Recordset
                        strSQL = "SELECT * FROM ODASMCouncilRateDue  ;"
                        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                        rsSAVE.AddNew
                        
                        '/ Obtain the Job Brief Details
                        
                        Set rsJOBBRIEF = New ADODB.Recordset
                        strJOBBRIEF = "SELECT * FROM ODASMJobBriefItems where JobBriefItemNo = '" & rsCONTROL!JobBriefItemNo & "'  ;"
                        rsJOBBRIEF.Open strJOBBRIEF, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                        If rsJOBBRIEF.EOF Or rsJOBBRIEF.BOF Then GoTo test
                        
                        '/ obtain the Payment mode based on the Council Selected
                        
                        Set rsCOUNCIL = New ADODB.Recordset
                        StrCOUNCIL = "SELECT * FROM ODASPCouncil WHERE CouncilCode = '" & rsCONTROL!CouncilCode & "' ;"
                        rsCOUNCIL.Open StrCOUNCIL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        If rsCOUNCIL.BOF Or rsCOUNCIL.EOF Then GoTo test
            
                        '/ obtain the first and the Last Day of the Year
                                    
                        Dim rsYEAR As ADODB.Recordset, stryear As String
                        Set rsYEAR = New ADODB.Recordset
                        stryear = "SELECT * FROM ODASPYear WHERE ODASPYear.currentYear = '" & CurrentYear & "' ;"
                        rsYEAR.Open stryear, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                        If rsYEAR.BOF Or rsYEAR.EOF Then Exit Sub
            
                        If rsCOUNCIL!UseCalendarYear = "Y" Then
                                rsSAVE!PaymentMode = "Y"
                                rsSAVE!Duration = 12
                                rsSAVE!StartDate = CDate(rsYEAR!StartDate)
                                rsSAVE!EndDate = CDate(rsYEAR!EndDate)
            
                        Else
                                rsSAVE!PaymentMode = "Q"
                                rsSAVE!Duration = 3
                                rsSAVE!StartDate = CDate(rsYEAR!StartDate)
                                rsSAVE!EndDate = DateAdd("M", CDbl(rsSAVE!Duration), CDate(rsSAVE!StartDate))
                                
                                If CDate(rsSAVE!EndDate) > CDate(dtLastDateOfYear) Then
                                        rsSAVE!EndDate = CDate(dtLastDateOfYear)
                                End If
                                            
                                rsSAVE!EndDate = DateAdd("D", -1, CDate(rsSAVE!EndDate))
                        End If
                        
                        loadCOUNCILRATES_batch
                        calcDUEDATE_batch
                        updateRATESCHEDULE_batch
                        
                        If rsSAVE!PaymentMode <> "Y" Then
                                createINSTALLMENTS
                        End If
test:
                    frmODASMEOY.ProgressBar1.Value = frmODASMEOY.ProgressBar1.Value + 1
                    rsCONTROL.MoveNext
            Loop
            
            frmODASMEOY.Visible = False

'rsCONTROL.Close
'rsJOBBRIEF.Close
'rsCOUNCIL.Close
'rsYEAR.Close
'
'Strcontrol = Empty
'strSQL = Empty
'strJOBBRIEF = Empty
'StrCOUNCIL = Empty
'stryear = Empty

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub saveRecord()
On Error GoTo err

    With frmODASMCouncilRates
    
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "Select * From ODASPYear Where CurrentYear = '" & Year(Date) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
        
        dtLastDateOfYear = rsFindRecord!EndDate
        
'        Do While CDate(.txtRateExpiryDate) <= CDate(.txtJBExpiryDate) And _
'              CDate(.txtRateExpiryDate) <= CDate(dtLastDateOfYear) And _
'              CDate(.txtRateStartDate) <= CDate(dtLastDateOfYear)
'
        Do While CDate(.txtRateExpiryDate) <= CDate(dtLastDateOfYear) And _
              CDate(.txtRateStartDate) <= CDate(dtLastDateOfYear)
              
                updateRATESCHEDULE
                UpdateJobBriefItems
                
                .txtRateStartDate.Text = DateAdd("d", 1, .txtRateExpiryDate.Text)
                .txtRateExpiryDate.Text = DateAdd("m", CDbl(.txtDuration.Text), CDate(.txtRateStartDate))
                                
                '/ Checking whether the last date is within the current Calendar year
                
                If CDate(.txtRateExpiryDate) > CDate(dtLastDateOfYear) Then
                        .txtRateExpiryDate.Text = CDate(dtLastDateOfYear)
                Else: .txtRateExpiryDate.Text = DateAdd("d", -1, CDate(.txtRateExpiryDate))
                End If

                
                calcDUEDATE
                .txtReferenceNo.Text = Empty
        Loop
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub createINSTALLMENTS()
On Error GoTo err

        dtStartDate = CDate(rsSAVE!StartDate)
        dtExpiryDate = CDate(rsSAVE!EndDate)
        intDURATION = CInt(rsSAVE!Duration)
        strPaymentMode = CStr(rsSAVE!PaymentMode)
        
        Do While CDate(dtExpiryDate) <= CDate(rsJOBBRIEF!expirydate) And _
              CDate(dtExpiryDate) <= CDate(dtLastDateOfYear) And _
              CDate(dtStartDate) <= CDate(dtLastDateOfYear)

                rsSAVE.AddNew
                
                rsSAVE!SiteNo = rsCONTROL!SiteNo
                rsSAVE!DatePrepared = Date
                rsSAVE!PreparedBY = CurrentUserName
                rsSAVE!CurrentYear = CurrentYear
                rsSAVE!paid = "N"
                rsSAVE!Requisitioned = "N"
                rsSAVE!Status = "RATES-PREPARED"
                rsSAVE!StartDate = DateAdd("d", 1, dtExpiryDate)
                rsSAVE!EndDate = DateAdd("m", CDbl(intDURATION), CDate(rsSAVE!StartDate))
                rsSAVE!Duration = intDURATION
                rsSAVE!PaymentMode = strPaymentMode
                rsSAVE!JobBriefItemNo = rsJOBBRIEF!JobBriefItemNo

                '/ Checking whether the last date is within the current Calendar year
                
                If CDate(rsSAVE!EndDate) > CDate(dtLastDateOfYear) Then
                        rsSAVE!EndDate = CDate(dtLastDateOfYear)
                Else: rsSAVE!EndDate = DateAdd("d", -1, CDate(rsSAVE!EndDate))
                End If

                loadCOUNCILRATES_batch

                calcDUEDATE_batch
                dtExpiryDate = rsSAVE!EndDate
                dtStartDate = CDate(rsSAVE!StartDate)

                rsSAVE.Update
        Loop

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub generateRATESCHEDULE()
On Error GoTo err
'
'        '/ Procedure to obtain the Last Day of The Year
'
'            obtainLASTDAY
'
'        '/ procedure to retrieve all the site that are active
'        '/
'
'        Set rsCONTROL = New ADODB.Recordset
'        strCONTROL = "SELECT * FROM ODASPPlotSite WHERE ACTIVE = 'Y' ;"
'        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
'
'
'        With rsCONTROL
'
'                '/ Procedure to retrieve the Job Brief Item Assigned to the
'                '/ Site
'
'                Set rsJOBBRIEF = New ADODB.Recordset
'                strJOBBRIEF = "SELECT * FROM ODASMJobBriefItems WHERE JobBriefItemNo = '" & rsCONTROL!JobBriefItemNo & "' ;"
'                rsJOBBRIEF.Open strJOBBRIEF, cnCOMMON, adOpenKeyset, adLockOptimistic
'
'                loadSTARTDATE
'
'                Do While CDate(rsJOBBRIEF!ExpiryDate) <= CDate(rsCONTROL!ExpiryDate) And _
'                CDate(.txtRateExpiryDate) <= CDate(dtLastDateOfYear) And _
'                CDate(.txtRateStartDate) <= CDate(dtLastDateOfYear)
'
'
'                updateRATESCHEDULE
'
'                .txtRateStartDate.Text = DateAdd("d", 1, .txtRateExpiryDate.Text)
'                .txtRateExpiryDate.Text = DateAdd("m", CDbl(.txtDuration.Text), CDate(.txtRateStartDate))
'
'                '/ Checking whether the last date is within the current Calendar year
'
'                If CDate(.txtRateExpiryDate) > CDate(dtLastDateOfYear) Then
'                        .txtRateExpiryDate.Text = CDate(dtLastDateOfYear)
'                Else: .txtRateExpiryDate.Text = DateAdd("d", -1, CDate(.txtRateExpiryDate))
'                End If
'
'
'                calcDUEDATE
'                .txtReferenceNo.Text = Empty
'        Loop
'    End With
'
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateRATESCHEDULE()
On Error GoTo err

    With frmODASMCouncilRates

            Set rsSAVE = New ADODB.Recordset
            
            If (bBillBoard = True Or bStreetSign = True) Then
                strSQL = "Select * from ODASMCouncilRateDue Where JobBriefItemNo= '" & .txtJobBriefItemNo.Text & "' and SiteNo = '" & .txtMast & "' and StartDate = '" & Format(.txtRateStartDate.Text, "MMMM dd,YYYY") & "' "
            Else
                strSQL = "Select * from ODASMCouncilRateDue Where JobBriefItemNo= '" & .txtJobBriefItemNo.Text & "' and SiteNo = '" & .txtSiteNo & "' and StartDate = '" & Format(.txtRateStartDate.Text, "MMMM dd,YYYY") & "' "
            End If
            
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.BOF Or rsSAVE.EOF Then
                rsSAVE.AddNew
                
                If (bBillBoard = True Or bStreetSign = True) Then
                    rsSAVE!SiteNo = .txtMast.Text
                    rsSAVE!BillBoard = "Y"
                    rsSAVE!Face = "N"
                Else
                    rsSAVE!SiteNo = .txtSiteNo.Text
                    rsSAVE!Face = "Y"
                    rsSAVE!BillBoard = "N"
                End If
                rsSAVE!JobBriefItemNo = .txtJobBriefItemNo.Text
                rsSAVE!StartDate = Format(.txtRateStartDate.Text, "MMMM dd,YYYY")
                rsSAVE!DatePrepared = Date
                rsSAVE!PreparedBY = CurrentUserName
                rsSAVE!CurrentYear = .txtCurrentYear.Text
                rsSAVE!paid = "N"
                rsSAVE!Requisitioned = "N"
            End If
                If (bBillBoard = True Or bStreetSign = True) Then
                    rsSAVE!BillBoard = "Y"
                    rsSAVE!Face = "N"
                Else
                    rsSAVE!Face = "Y"
                    rsSAVE!BillBoard = "N"
                End If
                rsSAVE!EndDate = Format(.txtRateExpiryDate.Text, "MMMM dd,YYYY")
                rsSAVE!AmountDue = CDbl(.txtAmount.Text)
                rsSAVE!DueDate = .txtRateDueDate.Text
                rsSAVE!PaymentMode = .cboPaymentMode.Text
                rsSAVE!Duration = CInt(.txtDuration.Text)
                rsSAVE!Status = "RATES-PREPARED"
                rsSAVE!Balance = CDbl(.txtAmount.Text)
                bSaveRECORD = False

            rsSAVE.Update
            .txtReferenceNo.Text = rsSAVE!ReferenceNo

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
Public Sub updateRATESCHEDULE_batch()
On Error GoTo err

    With frmODASMCouncilRates

            If baddRECORD = True Then
                    rsSAVE.AddNew
            End If
            
            rsSAVE!SiteNo = rsCONTROL!SiteNo
            rsSAVE!DatePrepared = Date
            rsSAVE!PreparedBY = CurrentUserName
            rsSAVE!CurrentYear = CurrentYear
            rsSAVE!paid = "N"
            rsSAVE!Requisitioned = "N"
            rsSAVE!Status = "RATES-PREPARED"
            rsSAVE!JobBriefItemNo = rsJOBBRIEF!JobBriefItemNo
                                           
            bSaveRECORD = False
            rsSAVE.Update

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
Public Sub updateSITE_batch()
On Error GoTo err

    With frmODASMCouncilRates

            
            
            Set rsSITE = New ADODB.Recordset
            
            strSQL = "Select * from ODASPPlotSite Where siteNo = '" & rsCONTROL!SiteNo & "' "
            rsSITE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSITE.BOF Or rsSITE.EOF Then Exit Sub
                        
            rsSITE!RateStatus = "REQUISITIONED"
            rsSITE!RateDue = CDbl(rsCONTROL!Amount)
            rsSITE!RateDueDate = CDate(rsCONTROL!DueDate)
                                
            bSaveRECORD = False

            rsSITE.Update
  End With

Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsSITE.CancelUpdate
            rsSITE.Requery
    Else
        UpdateErrorMessage
    End If

End Sub
Public Sub UpdateJobBriefItems()
On Error GoTo err
With frmODASMCouncilRates
    Set rsFIND = New ADODB.Recordset
    rsFIND.Open "Select * From ODASMJObBriefItems where JobBriefItemNo = '" & .txtJobBriefItemNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsFIND.EOF And rsFIND.BOF Then Exit Sub
        rsFIND!RatesComputed = "Y"
    rsFIND.Update
End With
Exit Sub
err:
ErrorMessage
End Sub

Public Sub updateSITE()
On Error GoTo err

    With frmODASMCouncilRates

            
            If .txtMediaCode.Text = "BIL" Then
                Set rsFIND = New ADODB.Recordset
                rsFIND.Open "Select * From ODASPPlotSite where MastNo = '" & .txtMast.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsFIND.EOF And rsFIND.BOF Then Exit Sub
                
                rsFIND.MoveFirst
                Do While rsFIND.EOF <> True
                Set rsSAVE = New ADODB.Recordset
                    strSQL = "Select * from ODASPPlotSite Where siteNo = '" & rsFIND!SiteNo & "' "
                    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        rsSAVE!RateStatus = .txtStatus.Text
                        rsSAVE!RateDue = CDbl(.txtAmount.Text)
                        rsSAVE!RateDueDate = CDate(.txtRateDueDate.Text)
                        
                        rsSAVE.Update
                    rsFIND.MoveNext
                Loop
            Else
                Set rsSAVE = New ADODB.Recordset
                strSQL = "Select * from ODASPPlotSite Where siteNo = '" & .txtSiteNo.Text & "' "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                    If rsSAVE.BOF Or rsSAVE.EOF Then Exit Sub
                    
                    rsSAVE!RateStatus = .txtStatus.Text
                    rsSAVE!RateDue = CDbl(.txtAmount.Text)
                    rsSAVE!RateDueDate = CDate(.txtRateDueDate.Text)
                                        
        
                    rsSAVE.Update
           End If
                    bSaveRECORD = False
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
Public Sub loadCOUNCILRATES_batch()
On Error GoTo err
                
                Dim rsCOUNCIL As ADODB.Recordset, StrCOUNCIL As String
                Set rsCOUNCIL = New ADODB.Recordset
                StrCOUNCIL = "SELECT * FROM ODASPLandRate R WHERE R.TownCode = '" & rsCONTROL!CouncilCode & "' and R.MediaCode = '" & rsCONTROL!MediaCode & "' and R.MediaSize = '" & rsCONTROL!MediaSize & "' and R.currentyear = '" & CurrentYear & "' and R.PaymentMode = '" & rsSAVE!PaymentMode & "';"
                rsCOUNCIL.Open StrCOUNCIL, cnCOMMON, adOpenKeyset, adLockOptimistic

                If rsCOUNCIL.EOF And rsCOUNCIL.BOF Then
                        rsSAVE!AmountDue = 0
                ElseIf IsNull(rsCOUNCIL) = True Then
                        rsSAVE!AmountDue = 0
                Else: rsSAVE!AmountDue = CDbl(rsCOUNCIL!Amount)
                End If
                
                rsSAVE!Balance = CDbl(rsSAVE!AmountDue)
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadCOUNCILRATES()
On Error GoTo err
        With frmODASMCouncilRates
                
                If .cboPaymentMode.Text = Empty Then Exit Sub
            
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASPLandRate WHERE TownCode = '" & .txtTownCode & "' and MediaCode = '" & .txtMediaCode.Text & "' and MediaSize = '" & .txtMediaSize & "' and currentyear = '" & .txtCurrentYear.Text & "' and PaymentMode = '" & .cboPaymentMode.Text & "';"
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                    If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                .txtAmount.Text = FormatNumber(rsCONTROL!Amount)
                
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadJOBBRIEFITEMS()
On Error GoTo err
        With frmODASMCouncilRates
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMJobBriefItems,ODASPMedia WHERE ODASMJobBriefItems.JobBriefItemNo = '" & .txtJobBriefItemNo & "' and ODASMJobBriefItems.MediaCode = ODASPMedia.MediaCode;"
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                    .txtJBStartDate.Text = rsCONTROL!CommencementDate
                    .txtJBExpiryDate.Text = rsCONTROL!expirydate
                    .txtJBDuration.Text = rsCONTROL!Duration
                    .txtJBDurationMode.Text = rsCONTROL!DurationMode
                    
                    If rsCONTROL!BillBoard = "Y" Then
                        bBillBoard = True
                    ElseIf rsCONTROL!StreetSign = "Y" Then
                        bStreetSign = True
                    Else
                        bBillBoard = False: bStreetSign = False
                    End If
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub ValidateRECORD()
On Error GoTo err
        With frmODASMCouncilRates
        
                bSaveRECORD = False
                          
                If CDbl(.txtAmount) <= 0 Then
                        MsgBox "The Council Rate Amount MUST be Greater Than Zero"
                        .txtAmount.SetFocus
                
                ElseIf .txtCurrentYear.Text <= Empty Then
                        MsgBox "The Current Year MUST NOT be Blank"
                        .txtCurrentYear.SetFocus
                        
                ElseIf .txtAccountNo.Text = Empty Then
                        MsgBox "The Account no entered is Invalid"
                        .txtAccountNo.SetFocus
                
                ElseIf .txtJobBriefItemNo.Text = Empty Then
                        MsgBox "A Job Brief Item Number is Required"
                        .txtJobBriefItemNo.SetFocus
                        
                ElseIf CDbl(.txtDuration.Text) <= 0 Then
                        MsgBox "The Duration MUST be Greater Than Zero"
                        .txtDuration.SetFocus
                
                ElseIf .txtMediaCode.Text = Empty Then
                        MsgBox "The Media Is Required "
                        .txtMediaCode.SetFocus
                
                ElseIf .txtMediaSize.Text = Empty Then
                        MsgBox "The Size of the Media is Required"
                        .txtMediaSize.SetFocus
                
                ElseIf .txtPlotNo.Text <= Empty Then
                        MsgBox "The Plot Number is Required"
                        .txtPlotNo.SetFocus
                
                ElseIf .txtSiteNo.Text <= Empty Then
                        MsgBox "The Site Number is Required"
                        .txtSiteNo.SetFocus
                
                ElseIf CDate(.txtRateDueDate.Text) <= Empty Then
                        MsgBox "The Date the Council Rate is Due is Manadatory"
                        .txtRateDueDate.SetFocus
                
                ElseIf CDate(.txtRateExpiryDate.Text) <= Empty Then
                        MsgBox "The Exppiry Date for the Council Rate is Required"
                        .txtRateExpiryDate.SetFocus
                        
                ElseIf CDate(.txtRateStartDate.Text) <= Empty Then
                        MsgBox "The Start Date for the Council Rate is Required"
                        .txtRateStartDate.SetFocus
                
                ElseIf CDate(.txtRateDueDate.Text) > CDate(.txtRateStartDate.Text) Then
                        MsgBox "The Due date MUST come before the Start Date for the Rate"
                        .txtRateDueDate.SetFocus
                Else
                        bSaveRECORD = True
                End If
        End With

Exit Sub

err:
    ErrorMessage
End Sub



Public Sub obtainLASTDAY()
On Error GoTo err
    
        Dim rsLASTDAY As ADODB.Recordset, strLASTDAY As String
        Set rsLASTDAY = New ADODB.Recordset
        
        strLASTDAY = "Select * from ODASPYear where currentYear = '" & CurrentYear & "'"
        rsLASTDAY.Open strLASTDAY, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsLASTDAY.BOF Or rsLASTDAY.EOF Then Exit Sub
        dtLastDateOfYear = rsLASTDAY!EndDate
    


Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calcDUEDATE_batch()
On Error GoTo err
        
        Dim rsDUEDATE As ADODB.Recordset, strDUEDATE As String
        Set rsDUEDATE = New ADODB.Recordset
        
        strDUEDATE = "Select * from ODASPDefault "
        rsDUEDATE.Open strDUEDATE, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsDUEDATE.BOF Or rsDUEDATE.EOF Then Exit Sub
        
        rsSAVE!DueDate = DateAdd("d", -1 * CDbl(rsDUEDATE!DefaultRateDays), CDate(rsSAVE!StartDate))
    
rsDUEDATE.Close
strDUEDATE = Empty

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calcDUEDATE()
On Error GoTo err
    
        Set rsCONTROL = New ADODB.Recordset
        
        strCONTROL = "Select * from ODASPDefault "
        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.BOF Or rsCONTROL.EOF Then Exit Sub
        
        With frmODASMCouncilRates
                .txtRateDueDate.Text = DateAdd("d", -1 * CDbl(rsCONTROL!DefaultRateDays), CDate(.txtRateStartDate))
        End With
    

Exit Sub

err:
    ErrorMessage
End Sub
