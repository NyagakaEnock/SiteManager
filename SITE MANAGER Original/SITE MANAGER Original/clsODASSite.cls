VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub saveRecord()
        updateRECORD
End Sub


Public Sub clearRECORD()
On Error GoTo err
        With frmODASMSiteRegistration
            .txtAcquisitionDate.Text = Empty
            .txtAcquisitionDate.Text = Date
            .txtAnnualRate.Text = 0
            .txtAnnualRent.Text = 0
            .txtExpiryDate.Text = Empty
            .txtLeaseDuration.Text = 0
            .txtLRNo.Text = Empty
            .txtPhysicalAddress.Text = Empty
            .txtPlotNo.Text = Empty
        End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub ClearPLOT()
On Error GoTo err
        With frmODASMSiteRegistration
            .txtAcquisitionDate.Text = Empty
            .txtAcquisitionDate.Text = Date
            .txtAnnualRate.Text = 0
            .txtAnnualRent.Text = 0
            .txtExpiryDate.Text = Empty
            .txtLeaseDuration.Text = 0
            .txtLRNo.Text = Empty
            .txtPhysicalAddress.Text = Empty
            .txtPlotNo.Text = Empty
            .txtLRNo.Text = Empty
            .txtZone.Text = Empty
            .DTPickerCommencementDate.Value = Date
            .DTPickerAcquisitionDate.Value = Date
            .DTPickerPropertyCommencementDate.Value = Date
            .DTPickerPropertyDateAssigned.Value = Date
        End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub enableOptions()
On Error GoTo err
        With frmODASMSiteRegistration
            .Frame7.Enabled = True
            .optCLIENT.Enabled = True
            .optEAGLE.Enabled = True
            .optNO.Enabled = True
            .optYES.Enabled = True
        End With
Exit Sub
err:
    ErrorMessage

End Sub
Public Sub enablePLOT()
On Error GoTo err
        With frmODASMSiteRegistration
        
            .txtAcquisitionDate.Locked = False
            .txtAcquisitionDate.Locked = True
            .txtAnnualRate.Locked = True
            .txtAnnualRent.Locked = True
            .txtMeterNo.Locked = False
            .cboCouncil.Locked = False
            .txtTownCity.Locked = False
            .txtExpiryDate.Locked = True
            .txtLeaseDuration.Locked = False
            .txtLandLordName.Locked = False
            .txtMobileNo.Locked = False
            .txtPostalAddress.Locked = False
            .txtemailAddress.Locked = False
            .txtLRNo.Locked = False
            .txtPhysicalAddress.Locked = False
            .txtPlotNo.Locked = True
            .txtCommencementDate.Locked = False
            .txtLRNo.Locked = False
            .txtZone.Locked = False
            .txtCouncil.Locked = True
            .txtLandLordNo.Locked = True
            .txtAcquiredBy.Locked = False
            .cboPaymentMode.Locked = False
            .DTPickerCommencementDate.Enabled = True
            .DTPickerAcquisitionDate.Enabled = True
            .UpDownLeasePeriod.Enabled = True
            .txtAnnualRentIncrement.Locked = True
            .txtAnnualRentIncrementType.Locked = False
            .txtRentVariationType.Locked = False
            .txtIncrementStartYear.Locked = False
            .txtIncrementFrequency.Locked = False
            .txtPaymentInterval.Locked = False
            .txtComments.Enabled = True
            .txtComments.Locked = False
            .chkWithLease.Enabled = True
            .chkWithLease.Value = vbChecked
            
        End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub disablePLOT()
On Error GoTo err
        With frmODASMSiteRegistration
            .txtAcquisitionDate.Locked = True
            .txtAcquisitionDate.Locked = True
            .txtAnnualRate.Locked = True
            .txtAnnualRent.Locked = True
            .txtExpiryDate.Locked = True
            .txtLeaseDuration.Locked = True
            .txtLRNo.Locked = True
            .txtPhysicalAddress.Locked = True
            .txtPlotNo.Locked = True
            .txtLRNo.Locked = True
        End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub clearMAST()
On Error GoTo err
    With frmODASMSiteRegistration
        .txtMastNo.Text = Empty
        .txtMastDetails.Text = Trim(.txtPhysicalAddress.Text)
        .txtMastAnnualRate.Text = 0
        .txtMastAnnualRent.Text = 0
        .txtMeterNo.Text = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub enableMAST()
On Error GoTo err
    With frmODASMSiteRegistration
        .txtMastNo.Locked = True
        .cboMedia.Locked = False
        .cboMedia.Enabled = True
        .txtMastAnnualRate.Locked = True
        .txtMastAnnualRent.Locked = False
        .txtMastDetails.Locked = False
        .txtMeterNo.Locked = False
        .txtTownCode.Locked = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub disableMAST()
On Error GoTo err
    With frmODASMSiteRegistration
        .txtMastNo.Locked = True
        .txtMastAnnualRate.Locked = True
        .txtMastDetails.Locked = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub clearSITE()
On Error GoTo err
    With frmODASMSiteRegistration
        .txtSiteNo.Text = Empty
        .txtSiteDetails.Text = Trim(.txtMastDetails.Text)
        .txtSiteAnnualRate.Text = 0
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub enableSITE()
On Error GoTo err
    With frmODASMSiteRegistration
        .txtRentVariationType.Locked = False
        .txtSiteNo.Locked = True
        .txtpaymentMode.Locked = True
        .txtSiteDetails.Locked = False
        .txtMastAnnualRate.Locked = True
        .chkSiteActive.Value = 1
        .txtAnnualRentIncrement.Locked = False
        .txtContractNo.Locked = True
        .txtSiteAnnualRate.Locked = False
        .UpDownLeasePeriod.Enabled = True
        .txtMeterNo.Enabled = True
        .txtZone.Locked = False
        .txtPaymentInterval.Locked = False
        .cboMediaSize.Locked = False
        .cboMedia.Locked = False
        .txtMastNo.Locked = True
        .txtTownCode.Locked = True
        .txtMastAnnualRent.Locked = False
        .txtIncrementFrequency.Locked = False
        .txtMastDetails.Locked = False
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub enableContract()
On Error GoTo err
    With frmODASMSiteRegistration
        .UpDownLeasePeriod.Enabled = True
        .DTPickerCommencementDate.Enabled = True
        .txtAnnualRentIncrementType = False
        .txtAnnualRentIncrement.Locked = False
        .txtPaymentInterval.Locked = False
        .txtIncrementFrequency.Locked = True
        .txtSiteDetails.Locked = False
        .cboPaymentMode.Enabled = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub disableSITE()
On Error GoTo err
    With frmODASMSiteRegistration
        .txtSiteNo.Locked = True
        .txtSiteDetails.Locked = True
        .txtSiteAnnualRate.Locked = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Function CouncilAccountNo()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASPPlot WHERE CouncilCode = '" & frmODASMSiteRegistration.cboCouncil.Text & "'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsCONTROL
                If .BOF = True And .EOF = True Then
                    CouncilAccountNo = frmODASMSiteRegistration.cboCouncil.Text & "-" & "001"
                ElseIf rsCONTROL.RecordCount <= 8 Then
                    CouncilAccountNo = frmODASMSiteRegistration.cboCouncil.Text & "-" & "00" & rsCONTROL.RecordCount + 1
                ElseIf rsCONTROL.RecordCount <= 9 And rsCONTROL.RecordCount < 100 Then
                    CouncilAccountNo = frmODASMSiteRegistration.cboCouncil.Text & "-" & "0" & rsCONTROL.RecordCount + 1
                Else
                    CouncilAccountNo = frmODASMSiteRegistration.cboCouncil.Text & "-" & rsCONTROL.RecordCount + 1
                End If
        End With
        
rsCONTROL.Close
strTRANS = ""

Exit Function
err:
    ErrorMessage
End Function
Public Sub GeneratePlotNo()
On Error GoTo err
With frmODASMSiteRegistration
        
        Dim rsLASTNO As ADODB.Recordset, strLAST As String
        
        Set rsLASTNO = New Recordset
        strLAST = "SELECT * FROM ODASPLastNumbers WHERE AutoPlotNo = 'Y'"
        rsLASTNO.Open strLAST, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        If rsLASTNO.BOF Or rsLASTNO.EOF = True Then Exit Sub
                
        Select Case Len(Trim(rsLASTNO!PlotNo))
        Case 1: .txtPlotNo.Text = Trim(rsLASTNO!PlotNoPrefix) + "00000" + Trim(rsLASTNO!PlotNo)
        Case 2: .txtPlotNo.Text = Trim(rsLASTNO!PlotNoPrefix) + "0000" + Trim(rsLASTNO!PlotNo)
        Case 3: .txtPlotNo.Text = Trim(rsLASTNO!PlotNoPrefix) + "000" + Trim(rsLASTNO!PlotNo)
        Case 4: .txtPlotNo.Text = Trim(rsLASTNO!PlotNoPrefix) + "00" + Trim(rsLASTNO!PlotNo)
        Case 5: .txtPlotNo.Text = Trim(rsLASTNO!PlotNoPrefix) + "0" + Trim(rsLASTNO!PlotNo)
        Case 6: .txtPlotNo.Text = Trim(rsLASTNO!PlotNoPrefix) + Trim(rsLASTNO!PlotNo)
        End Select
                
        rsLASTNO!PlotNo = rsLASTNO!PlotNo + 1
        rsLASTNO.Update
        rsLASTNO.Requery
        
        rsLASTNO.Close
        strLAST = ""


End With
        

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub GenerateMastNo()
On Error GoTo err

    With frmODASMSiteRegistration
            Set rsCONTROL = New Recordset
            
            strSQL = "SELECT count(*) as NoOfMasts  FROM ODASPPlotSite WHERE PlotNo = '" & frmODASMSiteRegistration.txtPlotNo.Text & "'"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
            If rsCONTROL.BOF Or rsCONTROL.EOF = True Then Exit Sub
            
            If IsNull(rsCONTROL!NoOfMasts) = True Then rsCONTROL!NoOfMasts = 0
                    .txtMastNo.Text = Trim(.txtPlotNo) + "-" + Trim(CDbl(rsCONTROL!NoOfMasts) + 1)
    End With
    
rsCONTROL.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub GenerateSiteNo()
On Error GoTo err

    With frmODASMSiteRegistration
        
        Set rsCONTROL = New Recordset
        strSQL = "SELECT * FROM ODASPPlotSite WHERE MastNo = '" & frmODASMSiteRegistration.txtMastNo.Text & "'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        If rsCONTROL.BOF Or rsCONTROL.EOF = True Then
                .txtSiteNo.Text = Trim(.txtMastNo) + "-1"
        Else: .txtSiteNo.Text = Trim(.txtMastNo) + "-" + Trim(CDbl(rsCONTROL.RecordCount) + 1)
        End If
    End With
        
rsCONTROL.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub updateRECORD()
On Error GoTo err
       With frmODASMSiteRegistration
            
                strSQL = "select * from ODASPPlot Where PlotNo = '" & .txtPlotNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        GeneratePlotNo
                        rsSAVE!CouncilAccountNo = CouncilAccountNo
                        rsSAVE!PlotNo = Trim(.txtPlotNo.Text)
                        rsSAVE!CreatedBy = CurrentUserName
                        rsSAVE!DateCreated = Date
                        rsSAVE!Created = "Y"
                        rsSAVE!Approved = "N"
                        rsSAVE!Authorized = "N"
                        rsSAVE!RentPaid = 0
                        rsSAVE!RatePaid = 0
                        rsSAVE!NoOfMasts = 0
                        rsSAVE!ContractStatus = "NOT-STARTED"
                        rsSAVE!ContractStatusDate = Date
                        rsSAVE!RateStatus = "NOT-PAID"
                        rsSAVE!RateStatusDate = Date
                        rsSAVE!RentStatus = "NOT-PAID"
                        rsSAVE!RentStatusDate = Date
                        rsSAVE!AcquisitionDate = Date
                End If
                
                rsSAVE!Status = .txtStatus.Text
                
                rsSAVE!AnnualRentIncrement = IIf(.txtAnnualRentIncrement = "", 0, .txtAnnualRentIncrement.Text)
                rsSAVE!PaymentMode = .cboPaymentMode
                rsSAVE!RentVariationType = .txtRentVariationType
                rsSAVE!AnnualRentIncrementType = .txtAnnualRentIncrementType
                rsSAVE!IncrementStartYear = .txtIncrementStartYear.Text
                rsSAVE!IncrementFrequency = .txtIncrementFrequency.Text
                rsSAVE!PaymentInterval = .txtPaymentInterval.Text
                rsSAVE!WithLease = IIf(.chkWithLease.Value = True, "N", "Y")
                
                rsSAVE!AcquisitionDate = CDate(.txtAcquisitionDate.Text)
                rsSAVE!MeterNo = .txtMeterNo.Text
                rsSAVE!PlotName = .txtLandLordName.Text
                rsSAVE!CouncilCode = .cboCouncil.Text
                rsSAVE!AccountNo = .txtLandLordNo.Text

                If .optYES.Value = True Then
                        rsSAVE!OnRoadReserve = "Y"
                Else: rsSAVE!OnRoadReserve = "N"
                End If
                
                rsSAVE!TownCode = .txtTownCode.Text
                rsSAVE!LeaseDuration = CDbl(.txtLeaseDuration.Text)
                rsSAVE!LRNo = .txtLRNo.Text
                rsSAVE!PhysicalLocation = Trim(.txtPhysicalAddress.Text)
                rsSAVE!CommencementDate = .txtCommencementDate.Text
                rsSAVE!expirydate = .txtExpiryDate.Text
                rsSAVE!AcquiredBy = .txtAcquiredBy.Text
                rsSAVE!Comments = .txtComments.Text
                
                '/ Procedure to Count the Number of Masts
                
                strCONTROL = "select  Count(PlotNo) as NoofMasts from ODASPPlotMast Where PlotNo = '" & .txtPlotNo & "';"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                        rsSAVE!NoOfMasts = 0
                ElseIf Trim(.txtMastNo.Text) <= "" Then
                        rsSAVE!NoOfMasts = 0
                ElseIf IsNull(rsCONTROL!NoOfMasts) = True Then
                        rsSAVE!NoOfMasts = 0
                Else: rsSAVE!NoOfMasts = rsCONTROL!NoOfMasts
                        
                        Set rsCONTROL = Nothing
                        strSQL = Empty
                End If
                
                '/ Procedure to Count the Number of Sites
                
                strCONTROL = "select  Count(SiteNo) as NoOfSites from ODASPPlotSite Where MastNo = '" & .txtMastNo & "';"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                        rsSAVE!NoofSites = 0
                        
                ElseIf Trim(.txtMastNo.Text) <= "" Then
                        rsSAVE!NoofSites = 0
                        
                ElseIf IsNull(rsCONTROL!NoofSites) = True Then
                        rsSAVE!NoofSites = 0
                Else
                        rsSAVE!NoofSites = rsCONTROL!NoofSites
                        
                        'Set rsCONTROL = Nothing
                        'strSQL = Empty
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
Public Sub updateANNUALRate()
On Error GoTo err
    With frmODASMSiteRegistration
                
                Set rsAnnualRate = New ADODB.Recordset
                rsAnnualRate.Open "SELECT * FROM ODASPPlot WHERE PlotNo = '" & .txtPlotNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                strSQL = "select * from ODASPPlotMast Where MastNo = '" & .txtMastNo.Text & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                                
                strCONTROL = "select  sum(ODASPPlotSite.Rates) as Totals from ODASPPlotSite Where PlotNo = '" & .txtPlotNo & "';"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Totals) = True Then
                        rsSAVE!AnnualRate = 0
                Else: rsSAVE!AnnualRate = rsCONTROL!Totals
                End If
                
                .txtMastAnnualRate.Text = rsSAVE!AnnualRate
                
                rsSAVE.Update
                rsSAVE.Requery
                
                
                strCONTROL = Empty
                strSQL = Empty
                rsCONTROL.Close
                rsSAVE.Close
                Set rsAnnualRate = Nothing
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateANNUALRent()
On Error GoTo err
    With frmODASMSiteRegistration
                
                strSQL = "select * from ODASPPlot Where PlotNo = '" & .txtPlotNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                                
                strCONTROL = "select  sum(ODASPPlotMast.AnnualRent) as Totals from ODASPPlotMast Where PlotNo = '" & .txtPlotNo & "';"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Totals) = True Then
                        rsSAVE!AnnualRent = 0
                Else: rsSAVE!AnnualRent = rsCONTROL!Totals
                End If
                
                .txtAnnualRent.Text = FormatNumber(rsSAVE!AnnualRent, 2)

                rsSAVE.Update
                rsSAVE.Requery

                strCONTROL = Empty
                strSQL = Empty
                rsCONTROL.Close
                rsSAVE.Close
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateNoOfMasts()
On Error GoTo err
    With frmODASMSiteRegistration
                
                strSQL = "select * from ODASPPlot Where PlotNo = '" & .txtPlotNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                                
                strCONTROL = "select  * from ODASPPlotMast Where PlotNo = '" & .txtPlotNo & "';"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                        rsSAVE!NoOfMasts = 0
                ElseIf IsNull(rsCONTROL.RecordCount) = True Then
                        rsSAVE!NoOfMasts = 0
                Else: rsSAVE!NoOfMasts = rsCONTROL.RecordCount
                End If
                
                rsSAVE.Update
                rsSAVE.Requery

                strCONTROL = Empty
                strSQL = Empty
                rsCONTROL.Close
                rsSAVE.Close
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateALLSites()
On Error GoTo err
    With frmODASMSiteRegistration
                
                strSQL = "select * from ODASPPlot Where PlotNo = '" & .txtPlotNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                                
                strCONTROL = "select  COUNT(plotNo) AS SiteCount from ODASPPlotSite Where PlotNo = '" & .txtPlotNo & "';"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                 If rsCONTROL.EOF Or rsCONTROL.BOF Then
                    rsSAVE!NoofSites = 0
                ElseIf IsNull(rsCONTROL!SiteCount) = True Then
                    rsSAVE!NoofSites = 0
                Else: rsSAVE!NoofSites = rsCONTROL!SiteCount
                End If
                
                .txtNoofSites.Text = rsSAVE!NoofSites
                
                strCONTROL = "select  COUNT(plotNo) AS MastCount from ODASPPlotMast Where PlotNo = '" & .txtPlotNo & "';"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                 If rsCONTROL.EOF Or rsCONTROL.BOF Then
                    rsSAVE!NoOfMasts = 0
                ElseIf IsNull(rsCONTROL!MastCount) = True Then
                    rsSAVE!NoOfMasts = 0
                Else: rsSAVE!NoOfMasts = rsCONTROL!MastCount
                End If

                .txtNoofMasts.Text = rsSAVE!NoOfMasts
                
                rsSAVE.Update
                rsSAVE.Requery

                strCONTROL = Empty
                strSQL = Empty
                rsCONTROL.Close
                rsSAVE.Close
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateNoOfSites()
On Error GoTo err
    With frmODASMSiteRegistration
                
                strSQL = "select * from ODASPPlotMast Where MastNo = '" & .txtMastNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                                
                strCONTROL = "select  * from ODASPPlotSite Where MastNo = '" & .txtMastNo & "';"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                        rsSAVE!NoofSites = 0
                Else: rsSAVE!NoofSites = rsCONTROL.RecordCount
                End If
                
                rsSAVE.Update
                rsSAVE.Requery

                strCONTROL = Empty
                strSQL = Empty
                rsCONTROL.Close
                rsSAVE.Close
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadRECORD()
On Error GoTo err
       With frmODASMSiteRegistration
            
                strSQL = "select * from ODASPPlot Where PlotNo = '" & .txtPlotNo & "';"
                Set rsFindRecord = New ADODB.Recordset
                rsFindRecord.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsFindRecord.EOF Or rsFindRecord.BOF Then
                        MsgBox "Record Not Found !"
                        Exit Sub
                End If
                
                If rsFindRecord!OnRoadReserve = "N" Then
                        .optNO.Value = True
                Else: .optYES.Value = True
                End If
                 .txtPlotNo.Text = rsFindRecord!PlotNo & ""
                .txtStatus.Text = rsFindRecord!Status
                .cboPaymentMode.Text = rsFindRecord!PaymentMode & ""
                .txtRentVariationType.Text = rsFindRecord!RentVariationType & ""
                .txtAnnualRentIncrementType.Text = rsFindRecord!AnnualRentIncrementType & ""
                If IsNull(rsFindRecord!IncrementStartYear) Then
                     .txtIncrementStartYear.Text = 2
                Else
                    .txtIncrementStartYear.Text = rsFindRecord!IncrementStartYear & ""
                End If
                
                If IsNull(rsFindRecord!IncrementFrequency) Then
                     .txtIncrementFrequency.Text = 1
                Else
                    .txtIncrementFrequency.Text = rsFindRecord!IncrementFrequency & ""
                End If
                
                If IsNull(rsFindRecord!PaymentInterval) Then
                     .txtPaymentInterval.Text = 1
                Else
                    .txtPaymentInterval.Text = rsFindRecord!PaymentInterval & ""
                End If
                
                If IsNull(rsFindRecord!NoOfMasts) = True Then
                        .txtNoofMasts.Text = 0
                Else: .txtNoofMasts.Text = FormatNumber(rsFindRecord!NoOfMasts, 0)
                End If
                
                If IsNull(rsFindRecord!NoofSites) = True Then
                        .txtNoofSites.Text = 0
                Else: .txtNoofSites.Text = FormatNumber(rsFindRecord!NoofSites, 0)
                End If


                If IsNull(rsFindRecord!AnnualRentIncrement) = True Then
                        .txtAnnualRentIncrement.Text = 0
                Else: .txtAnnualRentIncrement.Text = FormatNumber(rsFindRecord!AnnualRentIncrement, 2)
                End If

                If IsNull(rsFindRecord!AnnualRent) = True Then
                        .txtAnnualRent.Text = 0
                Else: .txtAnnualRent.Text = FormatNumber(rsFindRecord!AnnualRent, 2)
                End If
                
                If IsNull(rsFindRecord!AnnualRate) = True Then
                        .txtAnnualRate.Text = 0
                Else: .txtAnnualRate.Text = FormatNumber(rsFindRecord!AnnualRate, 2)
                End If
                
                .txtAcquiredBy.Text = rsFindRecord!AcquiredBy & ""
                .txtTownCode.Text = rsFindRecord!TownCode
                
                If IsDate(rsFindRecord!AcquisitionDate) = True Then .txtAcquisitionDate.Text = rsFindRecord!AcquisitionDate & ""
                
                If IsNumeric(rsFindRecord!LeaseDuration) = False Then
                        .txtLeaseDuration.Text = 0
                        .UpDownLeasePeriod.Value = 1
                Else: .txtLeaseDuration.Text = rsFindRecord!LeaseDuration & ""
                        If rsFindRecord!LeaseDuration = 0 Then
                            .UpDownLeasePeriod.Value = 1
                        Else
                            .UpDownLeasePeriod.Value = rsFindRecord!LeaseDuration
                        End If
                End If
                
                .chkWithLease.Value = IIf(rsFindRecord!WithLease = "Y", True, False)
                
                .txtLRNo.Text = rsFindRecord!LRNo & ""
                .cboCouncil.Text = rsFindRecord!CouncilCode & ""
                
                .txtPhysicalAddress.Text = rsFindRecord!PhysicalLocation & ""
                
                If IsDate(rsFindRecord!CommencementDate) = True Then .txtCommencementDate.Text = rsFindRecord!CommencementDate & ""
                
                If IsDate(rsFindRecord!expirydate) = True Then .txtExpiryDate.Text = rsFindRecord!expirydate & ""
                
                
                .txtLandLordNo.Text = rsFindRecord!AccountNo & ""
                
                .txtComments.Text = rsFindRecord!Comments & ""
               
                Set rsFindRecord = Nothing
                strSQL = Empty
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub saveMAST()
On Error GoTo err
       With frmODASMSiteRegistration
           If .txtMastNo.Text = Empty And NewRecord Then
                        GenerateMastNo
           End If
                strSQL = "select * from ODASPPlotMast Where MastNo = '" & .txtMastNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!MastNo = Trim(.txtMastNo.Text)
                        rsSAVE!CreatedBy = CurrentUserName
                        rsSAVE!DateCreated = Date
                        rsSAVE!Created = "Y"
                        rsSAVE!Approved = "N"
                        rsSAVE!Authorized = "N"
                        rsSAVE!LeasePrepared = "N"
                        rsSAVE!RentPaid = 0
                        rsSAVE!RatePaid = 0
                        rsSAVE!AllocationDate = Date
                        rsSAVE!NoofSites = 0
                    
                End If
                
                rsSAVE!Status = .txtStatus.Text
                rsSAVE!MeterNo = .txtMeterNo.Text
                rsSAVE!PlotNo = Trim(.txtPlotNo.Text)
                rsSAVE!MastDetails = .txtMastDetails.Text
                
                If .optCLIENT.Value = True Then
                        rsSAVE!OwenedByClient = "Y"
                Else: rsSAVE!OwenedByClient = "N"
                End If
                
                rsSAVE!AnnualRent = CDbl(.txtMastAnnualRent.Text)
                rsSAVE!AnnualRate = CDbl(.txtMastAnnualRate.Text)
                rsSAVE!LeaseDuration = CDbl(.txtLeaseDuration.Text)
                rsSAVE!CommencementDate = .txtCommencementDate.Text
                rsSAVE!expirydate = .txtExpiryDate.Text
                rsSAVE!TypeOfMast = .cboMedia
                rsSAVE!MediaSize = .cboMediaSize
                
                rsSAVE.Update
                rsSAVE.Requery
                
                Set rsSAVE = Nothing
                strSQL = Empty
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub DeleteMast()
On Error GoTo err
       With frmODASMSiteRegistration
                
                strSQL = "Delete from ODASPPlotMast Where MastNo = '" & .txtMastNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Set rsSAVE = Nothing
                strSQL = Empty
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub saveSITE()
On Error GoTo err
       With frmODASMSiteRegistration
                
                If Len(Trim(.txtSiteNo.Text)) <= 3 And NewRecord = True Then GenerateSiteNo
                
                strSQL = "select * from ODASPPlotSite Where SiteNo = '" & .txtSiteNo.Text & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!SiteNo = Trim(.txtSiteNo.Text)
                        rsSAVE!CreatedBy = CurrentUserName
                        rsSAVE!DateCreated = Date
                        rsSAVE!Created = "Y"
                        rsSAVE!PropertiesAssigned = "N"
                        rsSAVE!Approved = "N"
                        rsSAVE!Authorized = "N"
                        rsSAVE!RatePaid = 0
                        rsSAVE!AllocationDate = Date
                        rsSAVE!JobBriefNo = Empty
                        rsSAVE!Active = "Y"
                End If
                
                rsSAVE!Status = .txtStatus.Text
                rsSAVE!MeterNo = .txtMeterNo.Text
                rsSAVE!HighwayZone = .txtZone.Text
                rsSAVE!MastNo = Trim(.txtMastNo.Text)
                rsSAVE!PlotNo = Trim(.txtPlotNo.Text)
                rsSAVE!SiteDetails = .txtSiteDetails.Text
                rsSAVE!MediaSize = .cboMediaSize.Text
                rsSAVE!MediaCode = .cboMedia.Text
                
                If .chkSiteActive.Value = 1 Then
                    rsSAVE!Active = "Y"
                Else: rsSAVE!Active = "N"
                End If
                
                rsSAVE!Rates = CDbl(.txtSiteAnnualRate.Text) & ""
                
                '/ Save Mast Related Data
                
                rsSAVE!StartDate = .txtCommencementDate
                rsSAVE!expirydate = .txtExpiryDate
                
                rsSAVE.Update
                rsSAVE.Requery
                
                Set rsSAVE = Nothing
                strSQL = Empty
        End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub loadMAST(Optional strMastNo As String = "")
On Error GoTo err
        With frmODASMSiteRegistration
                
                If strMastNo = "" Then
                    strSQL = "select * from ODASPPlotMast Where PlotNo = '" & .txtPlotNo & "';"
                Else
                    strSQL = "select * from ODASPPlotMast Where MastNo = '" & strMastNo & "';"
                End If
                
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                        Exit Sub
                End If
                
                .txtMastDetails.Text = rsCONTROL!MastDetails
                .txtMastNo.Text = rsCONTROL!MastNo
                
                If IsNull(rsCONTROL!AnnualRent) = True Then
                        .txtMastAnnualRent.Text = 0
                Else: .txtMastAnnualRent.Text = FormatNumber(rsCONTROL!AnnualRent, 0)
                End If
                
                If IsNull(rsCONTROL!AnnualRate) = True Then
                        .txtMastAnnualRate.Text = 0
                Else: .txtMastAnnualRate.Text = FormatNumber(rsCONTROL!AnnualRate, 2)
                End If
                
                .txtMeterNo.Text = rsCONTROL!MeterNo
                .cboMedia.Text = rsCONTROL!TypeOfMast & ""
                .cboMediaSize.Text = rsCONTROL!MediaSize & ""
                
                If rsCONTROL!OwenedByClient = "Y" Then
                        .optCLIENT.Value = True
                Else: .optEAGLE.Value = True
                End If
                
                Set rsCONTROL = Nothing
                strSQL = Empty
End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadSITE(Optional strSiteNo As String = "")
On Error GoTo err
        With frmODASMSiteRegistration
                
                If strSiteNo = "" Then
                        strSQL = "select * from ODASPPlotSite Where MastNo = '" & .txtMastNo.Text & "';"
                Else: strSQL = "select * from ODASPPlotSite Where SiteNo = '" & strSiteNo & "';"
                End If
                
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                        .txtSiteAnnualRate.Text = 0
                        
                        Exit Sub
                End If
                
                .txtSiteNo.Text = rsCONTROL!SiteNo
                .txtSiteDetails.Text = rsCONTROL!SiteDetails
                .txtSiteAnnualRate.Text = rsCONTROL!Rates & ""
                .cboMediaSize.Text = rsCONTROL!MediaSize & ""
                .cboMedia.Text = rsCONTROL!MediaCode & ""
                .txtZone.Text = rsCONTROL!HighwayZone & ""
                
                Set rsCONTROL = Nothing
                strSQL = Empty
End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub LoadContract(Optional strContractNo As String = "")
On Error GoTo err
With frmODASMSiteRegistration
                
                If strContractNo = "" Then
                        strSQL = "select * from ODASMLeaseAgreement Where ContractNo = '" & .txtContractNo.Text & "';"
                Else: strSQL = "select * from ODASMLeaseAgreement Where ContractNo = '" & strContractNo & "';"
                End If
                
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                        .txtCommencementDate.Text = ""
                        .txtLeaseDuration.Text = ""
                        .txtCommencementDate.Text = ""
                        .txtExpiryDate.Text = ""
                        .cboPaymentMode.Text = ""
                        .txtpaymentMode.Text = ""
                        .txtIncrementFrequency.Text = 1
                        .txtPaymentInterval.Text = 1
                        .txtIncrementStartYear.Text = 2
                        .txtAnnualRentIncrement.Text = 0
                        .txtAnnualRentIncrementType.Text = "N"
                        .txtContractNo.Text = ""
                Else
                        .txtLeaseDuration.Text = rsCONTROL!LeaseDuration & ""
                        .txtCommencementDate.Text = rsCONTROL!CommencementDate & ""
                        .txtExpiryDate.Text = rsCONTROL!expirydate & ""
                        .cboPaymentMode.Text = rsCONTROL!PaymentMode & ""
                        .txtpaymentMode.Text = rsCONTROL!PaymentMode & ""
                        .txtIncrementFrequency.Text = rsCONTROL!IncrementFrequency & ""
                        .txtPaymentInterval.Text = rsCONTROL!PaymentInterval & ""
                        .txtIncrementStartYear.Text = rsCONTROL!IncrementStartYear & ""
                        .txtAnnualRentIncrement.Text = rsCONTROL!AnnualRentIncrement & ""
                        .txtAnnualRentIncrementType.Text = rsCONTROL!AnnualRentIncrementType & ""
                        .txtContractNo.Text = rsCONTROL!ContractNo
                
                End If
End With

Exit Sub

err:
    ErrorMessage
End Sub


Public Function ValidDateRecord()
With frmODASMSiteRegistration
    ValidDateRecord = False
    If .txtLandLordName.Text = Empty Then
        MsgBox "Name Of site Required...!"
        .txtLandLordName.SetFocus
    ElseIf .txtCouncil.Text = Empty Then
        MsgBox "The Council is Required ...."
        .txtCouncil.SetFocus
    ElseIf .txtAcquisitionDate.Text = Empty Then
        MsgBox "The Acquisition Date is Required ...."
        .txtAcquisitionDate.SetFocus
    ElseIf DateDiff("D", .txtAcquisitionDate, Date) < 0 Then
        MsgBox "The Acquistion Date cannot be in the Future"
        .txtAcquisitionDate.SetFocus
    
    ElseIf .txtExpiryDate.Text = Empty Then
        MsgBox "The Expiry Date is Necessary ..........."
        .txtExpiryDate.SetFocus
    
    ElseIf .txtCommencementDate.Text = Empty Then
        MsgBox "The Commencement date of the Site is Required ......."
        .txtCommencementDate.SetFocus
    
'    ElseIf DateDiff("D", .txtCommencementDate, Date) < 0 Then
'        MsgBox "The Commencement Date cannot be in the Future"
'        .txtCommencementDate.SetFocus
    
    ElseIf CDbl(.txtLeaseDuration.Text) < 0 Then
        MsgBox "The Duration of the Lease is Required .........."
        .txtLeaseDuration.SetFocus
    
    ElseIf .txtLRNo.Text <= "" Then
        MsgBox "The LR Number is Required .............."
        .txtLRNo.SetFocus
    
    ElseIf .txtPhysicalAddress.Text <= "" Then
        MsgBox "The Physiccal Address of the Site is Necessary ......"
        .txtPhysicalAddress.SetFocus
    
        
    ElseIf .txtTownCode.Text <= "" Then
        MsgBox "The Town Code is Required ................"
            .txtTownCode.SetFocus
    
    ElseIf .txtMastAnnualRate.Text <= Empty Then
            MsgBox "The Annual Rate for the BillBoard is Required......"
            .txtMastAnnualRate.SetFocus
    
    
    ElseIf .txtMastAnnualRent.Text <= Empty Then
            MsgBox "The Annual Rent for the MAst is Required........"
            .txtMastAnnualRent.SetFocus
    
    ElseIf .txtZone.Text = Empty Then
            MsgBox "The Highway Zone is Required"
            .txtZone.SetFocus
    
    ElseIf .txtLeaseDuration.Text = Empty Then
            MsgBox "The duration of the Lease is Required ............."
            .txtLeaseDuration.SetFocus
    
    Else
            ValidDateRecord = True
    End If

End With
End Function

Public Function validFACE()
With frmODASMSiteRegistration
    validFACE = False
    
    If .txtLandLordName.Text = Empty Then
        MsgBox "Name Of LandLord Required...........................!"
        .txtLandLordName.SetFocus
    
    ElseIf .txtAnnualRentIncrementType.Text <> "N" And .txtAnnualRentIncrementType.Text <> "P" And .txtAnnualRentIncrementType.Text <> "A" Then
            MsgBox "The Annual Increment Can be (P)ercent, (A)mount or (N)one ................"
            .txtAnnualRentIncrementType.SetFocus
            
    ElseIf .txtLRNo.Text <= "" And .optNO.Value = True Then
        MsgBox "The LR No is Required ......................"
        .txtLRNo.SetFocus
     
    ElseIf IIf(.txtAnnualRentIncrement.Text = "", 0, .txtAnnualRentIncrement.Text) <= 0 And (.txtAnnualRentIncrementType.Text <> "N") Then
        MsgBox "You Must Indicate the Basis of the Annual Rent Increment as either P or Amount.........."
        .txtAnnualRentIncrement.SetFocus
    
    ElseIf .cboPaymentMode.Text <= "" And .optNO.Value = True Then
        MsgBox "The Payment Mode cannot be Left Blank, there is need to Compute the Installment Due...."
        .cboPaymentMode.SetFocus
        
    ElseIf .txtPostalAddress.Text <= Empty Then
        MsgBox "The Postal Address is Required ................................."
        .txtPostalAddress.SetFocus
    
    ElseIf .txtMobileNo.Text <= Empty Then
        MsgBox "The Mobile Phone Number for the Contact is Required ..........."
        .txtMobileNo.SetFocus
    
    ElseIf .txtemailAddress.Text <= Empty Then
        MsgBox "The Email address of the Contact is Required ..................."
        .txtemailAddress.SetFocus
        
    ElseIf .txtCouncil.Text = Empty Then
        MsgBox "The Council is Required ........................................"
        .txtCouncil.SetFocus
    
    ElseIf .cboMediaSize.Text = Empty Then
        MsgBox "The Media Size is Required ......................................"
        .cboMediaSize.SetFocus
    
    ElseIf .cboMedia.Text = Empty Then
        MsgBox "The Media type is Required ......................................"
        .cboMedia.SetFocus
    
    ElseIf .txtAcquisitionDate.Text = Empty Then
        MsgBox "The Acquisition Date is Required ................................"
        .txtAcquisitionDate.SetFocus
    
    ElseIf DateDiff("D", .txtAcquisitionDate, Date) < 0 Then
        MsgBox "The Acquistion Date cannot be in the Future......................."
        .txtAcquisitionDate.SetFocus
    
    ElseIf IsDate(.txtExpiryDate.Text) <> True Then
        MsgBox "The Expiry Date is Necessary ..........."
        .txtExpiryDate.SetFocus
    
    ElseIf .txtCommencementDate.Text = Empty Then
        MsgBox "The Commencement date of the Site is Required ......................"
        .txtCommencementDate.SetFocus
    ElseIf DateDiff("D", .txtCommencementDate, .txtExpiryDate) < 0 Then
        MsgBox "The Commencement Date cannot be greater than the expiry date......................."
        .txtCommencementDate.SetFocus
    
    ElseIf CDbl(.txtLeaseDuration.Text) <= 0 Then
        MsgBox "The Duration of the Lease is Required ..................................."
        .txtLeaseDuration.SetFocus
    
    ElseIf .txtLRNo.Text <= "" Then
        MsgBox "The LR Number is Required ..............................................."
        .txtLRNo.SetFocus
    
    ElseIf .txtPhysicalAddress.Text <= "" Then
        MsgBox "The Physiccal Address of the Site is Necessary .........................."
        .txtPhysicalAddress.SetFocus
    
        
    ElseIf .optCLIENT.Value = False And .optEAGLE.Value = False Then
        MsgBox "The specification of the Structure / Mast Ownership is Required ..............."
        .optEAGLE.SetFocus
        
    ElseIf .optEAGLE.Value = True And .txtMastAnnualRent.Text = "0" And .optNO.Value = True Then
        MsgBox "The Annual Rent for the Mast / Structure cannot be ZERO!..............."
        .txtMastAnnualRent.SetFocus
    
    ElseIf .optCLIENT.Value = True And .txtMastAnnualRent.Text > "0" Then
        MsgBox "The Annual Rent is not Payable when the Mast / Structure is Owned by the Client!.....Set it to ZERO."
        .txtMastAnnualRent.SetFocus
        
    ElseIf .optNO.Value = False And .optYES.Value = False Then
        MsgBox "The specification of the plot (whether on Road Reserve or not) is Required ..............."
        .optNO.Enabled = True
        
    ElseIf .optYES.Value = False And .txtMastAnnualRent.Text = "0" And .optEAGLE.Value = True Then
        MsgBox "The Annual rent For  the Mast / Structure cannot be ZERO!!!........"
        .txtMastAnnualRent.SetFocus
        
    ElseIf .txtTownCode.Text <= "" Then
        MsgBox "The Town Code is Required ........................................................"
            .txtTownCode.SetFocus
    
    ElseIf .txtMastAnnualRate.Text <= Empty Then
            MsgBox "The Annual Rate for the BillBoard is Required.................................."
            .txtMastAnnualRate.SetFocus
    
            
    ElseIf .txtSiteDetails.Text = Empty Then
            MsgBox "The Exact Location of the Mast is Required...................................."
            .txtSiteDetails.SetFocus
            
    ElseIf .txtSiteAnnualRate.Text <= Empty Then
            MsgBox "The Annual Rate for the Site is Required......"
            .txtSiteAnnualRate.SetFocus
    
    ElseIf .txtMastAnnualRent.Text <= Empty Then
            MsgBox "The Annual Rent for the MAst is Required........"
            .txtMastAnnualRent.SetFocus
    
    ElseIf .txtMastDetails.Text = Empty Then
            MsgBox "The Exact Location of the Site is Required................................."
            .txtMastDetails.SetFocus
    
    ElseIf .txtMeterNo.Text <= Empty Then
            MsgBox "The Meter No is Required .................................................."
            .txtMeterNo.SetFocus
            
    ElseIf .txtZone.Text = Empty Then
            MsgBox "The Highway Zone is Required ............................................"
            .txtZone.SetFocus
    ElseIf .txtIncrementStartYear < 0 Then
            MsgBox "Invalid Increment Start Installment .................................................."
            .txtIncrementStartYear.SetFocus
    
    ElseIf .txtIncrementFrequency < 0 Then
            MsgBox "Invalid Increment Frequency .................................................."
            .txtIncrementFrequency.SetFocus
    
    Else
    
            validFACE = True
    End If

End With
End Function

Public Function validMAST()
With frmODASMSiteRegistration
    validMAST = False
    If .txtCouncil.Text = Empty Then
        MsgBox "The Council is Required ...."
        .txtCouncil.SetFocus
    ElseIf .txtAcquisitionDate.Text = Empty Then
        MsgBox "The Acquisition Date is Required ...."
        .txtAcquisitionDate.SetFocus
    ElseIf DateDiff("D", .txtAcquisitionDate, Date) < 0 Then
        MsgBox "The Acquistion Date cannot be in the Future"
        .txtAcquisitionDate.SetFocus
    
    ElseIf .txtExpiryDate.Text = Empty Then
        MsgBox "The Expiry Date is Necessary ..........."
        .txtExpiryDate.SetFocus
    
    ElseIf .txtCommencementDate.Text = Empty Then
        MsgBox "The Commencement date of the Site is Required ......."
        .txtCommencementDate.SetFocus
    
'    ElseIf DateDiff("D", .txtCommencementDate, Date) < 0 Then
'        MsgBox "The Commencement Date cannot be in the Future"
'        .txtCommencementDate.SetFocus
    
    ElseIf CDbl(.txtLeaseDuration.Text) < 0 Then
        MsgBox "The Duration of the Lease is Required .........."
        .txtLeaseDuration.SetFocus
    
    ElseIf .txtLRNo.Text <= "" Then
        MsgBox "The LR Number is Required .............."
        .txtLRNo.SetFocus
    
    ElseIf .txtPhysicalAddress.Text <= "" Then
        MsgBox "The Physiccal Address of the Site is Necessary ......"
        .txtPhysicalAddress.SetFocus
    
    ElseIf .optCLIENT.Value = False And .optEAGLE.Value = False Then
        MsgBox "The specification of the Structure / Mast Ownership is Required ..............."
        .optEAGLE.SetFocus
        
    ElseIf .optEAGLE.Value = True And .txtMastAnnualRent.Text = "0" And .optNO.Value = True Then
        MsgBox "The Annual Rent for the Mast / Structure cannot be ZERO!..............."
        .txtMastAnnualRent.SetFocus
    
    ElseIf .optCLIENT.Value = True And .txtMastAnnualRent.Text > "0" Then
        MsgBox "The Annual Rent is not Payable when the Mast / Structure is Owned by the Client!.....Set it to ZERO."
        .txtMastAnnualRent.SetFocus
        
    ElseIf .optNO.Value = False And .optYES.Value = False Then
        MsgBox "The specification of the plot (whether on Road Reserve or not) is Required ..............."
        .optNO.Enabled = True
        
    ElseIf .optYES.Value = False And .txtMastAnnualRent.Text = "0" And .optEAGLE.Value = True Then
        MsgBox "The Annual rent For  the Mast / Structure cannot be ZERO!!!........"
        .txtMastAnnualRent.SetFocus
        
    ElseIf .txtTownCode.Text <= "" Then
        MsgBox "The Town Code is Required ................"
            .txtTownCode.SetFocus
    
    ElseIf .txtMastAnnualRate.Text <= Empty Then
            MsgBox "The Annual Rate for the BillBoard is Required......"
            .txtMastAnnualRate.SetFocus
    
    
    ElseIf .txtMastAnnualRent.Text <= Empty Then
            MsgBox "The Annual Rent for the MAst is Required........"
            .txtMastAnnualRent.SetFocus
    
    ElseIf .txtMastDetails.Text = Empty Then
            MsgBox "The Exact Location of the Site is Required"
            .txtMastDetails.SetFocus
            
    Else
            validMAST = True
    End If

End With
End Function

Public Sub loadDEFAULTS()
On Error GoTo err
        With frmODASMSiteRegistration
                .txtAcquiredBy.Text = CurrentUserName
                .txtZone.Text = "A"
                .txtAcquisitionDate.Text = Date
                .txtMastAnnualRent.Text = 0
                .txtMastAnnualRate.Text = 0
                .txtAnnualRate.Text = 0
                .txtAnnualRent.Text = 0
                .txtCommencementDate.Text = Date
                .txtLeaseDuration.Text = 5
                .txtAcquiredBy.Text = CurrentUserName
                .txtStatus.Text = "SITE-ACQUIRED"
                .txtExpiryDate = DateAdd("YYYY", CDbl(.txtLeaseDuration), .txtCommencementDate.Text)
                .txtExpiryDate = DateAdd("Y", -1, .txtExpiryDate)
        End With
Exit Sub

err:
    ErrorMessage
End Sub

