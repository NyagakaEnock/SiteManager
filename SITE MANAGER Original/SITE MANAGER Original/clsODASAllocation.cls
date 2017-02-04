VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub loadDEFAULTS()
On Error GoTo err
    With Screen.ActiveForm
        .txtAmountDue.Text = 0
        .txtCommencementDate.Text = Date
        .DTPickerCommencementDate.Value = Date
        .txtIntervalCode.Text = 1
        .txtInterval.Text = "M"
        .txtAgreementDate.Text = Date
        .DTPickerAgreementDate.Value = Date
        .txtExpiryDate.Text = DateAdd("M", CDbl(.txtIntervalCode), Date)
    End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub loadRECORD()
On Error GoTo err
    With Screen.ActiveForm
        
            Set rsCONTROL = New ADODB.Recordset
            rsCONTROL.Open "SELECT * FROM ODASMLeaseAgreement, ODASPaCCOUNT  WHERE ODASMLeaseAgreement.ContractNo = '" & .txtContractNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
            .txtContractNo.Text = rsCONTROL!ContractNo
            .txtPlotNo.Text = rsCONTROL!PlotNo
            .txtLandLordNo.Text = rsCONTROL!AccountNo
            If IsDate(rsCONTROL!AgreementDate) Then .txtAgreementDate.Text = rsCONTROL!AgreementDate & ""
            .txtNames.Text = rsCONTROL!CompanyName & ""
            .txtSignedBy.Text = rsCONTROL!SignedBy & ""
            .txtWitnessCoy.Text = rsCONTROL!WitnessCoy & ""
            .txtWitnessLandLord.Text = rsCONTROL!WitnessLandLord
            
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub saveRecord()
On Error GoTo err
        With frmODASMAllocation
            k = 0: j = .ListView3.ListItems.Count
                For i = 1 To j
                    If .ListView3.ListItems(i).Checked = True Then
                    k = k + 1
                    End If
                Next i
        End With
        If k = 0 Then
                MsgBox ("Please select one or more BillBoards to Lease!"), vbCritical + vbOKOnly, "Lease Preparation"
        
        Else
        
        Set rsSAVE = New ADODB.Recordset
            rsSAVE.Open "SELECT * FROM ODASMLeaseAgreement  WHERE ContractNo = '" & frmODASMAllocation.txtContractNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
                    
            With frmODASMAllocation
                    If rsSAVE.EOF And rsSAVE.BOF Then
                            rsSAVE.AddNew
                            GenerateContractNo
                            LoadDEFAULT
                            rsSAVE!ContractNo = .txtContractNo.Text
                            rsSAVE!CompanyCode = "MAG"
                            rsSAVE!Preparedby = CurrentUserName
                            rsSAVE!dateprepared = Date
                            rsSAVE!Renewal = 1
                            rsSAVE!Status = "DRAFT AGREEMENT"
                              If rsDEFAULT!AutoApproval <> "Y" Then
                                  rsSAVE!Approved = "N"
                              Else:
                                  rsSAVE!Approved = "Y"
                              End If
                            rsSAVE!Authorized = "N"
                     End If
                    
                    If .chkDeallocate.Value = 1 Then
                            rsSAVE!AsSigned = "N"
                        Else: rsSAVE!AsSigned = "Y"
                    
                    End If
                    
                    rsSAVE!PaymentMode = .cboPaymentMode.Text
                    rsSAVE!PlotNo = .txtPlotNo.Text
                    rsSAVE!AgreementDate = .txtAgreementDate
                    rsSAVE!AccountNo = .txtLandLordNo.Text
                    rsSAVE!NoOfBillBoards = k
                    rsSAVE!WitnessLandLord = .txtWitnessLandLord.Text
                    If .chkYes.Value = 1 Then
                        rsSAVE!IncreamentAnnualRent = "Y"
                        rsSAVE!PercentageIncreament = .txtPercentage.Text
                    Else
                        rsSAVE!IncreamentAnnualRent = "N"
                        rsSAVE!PercentageIncreament = 0
                    End If
                    If .OptStandard.Value = True Then
                        rsSAVE!Standard = "Y"
                    Else:
                        rsSAVE!Standard = "N"
                    End If
                    
                    
                    rsSAVE.Update
            End With

        End With
End If
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub GenerateContractNo()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASPLastNumbers WHERE AutoContractNo = 'Y'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsCONTROL
                If .BOF Or .EOF = True Then Exit Sub
                        Screen.ActiveForm.txtContractNo.Text = !ContractNo & ""
                
                Select Case Len(Trim(Screen.ActiveForm.txtContractNo))
                        Case 1: Screen.ActiveForm.txtContractNo.Text = Trim(!ContractNoPrefix) + "00000" + Trim(Screen.ActiveForm.txtContractNo)
                        Case 2: Screen.ActiveForm.txtContractNo.Text = Trim(!ContractNoPrefix) + "0000" + Trim(Screen.ActiveForm.txtContractNo)
                        Case 3: Screen.ActiveForm.txtContractNo.Text = Trim(!ContractNoPrefix) + "000" + Trim(Screen.ActiveForm.txtContractNo)
                        Case 4: Screen.ActiveForm.txtContractNo.Text = Trim(!ContractNoPrefix) + "00" + Trim(Screen.ActiveForm.txtContractNo)
                        Case 5: Screen.ActiveForm.txtContractNo.Text = Trim(!ContractNoPrefix) + "0" + Trim(Screen.ActiveForm.txtContractNo)
                        Case 6: Screen.ActiveForm.txtContractNo.Text = Trim(!ContractNoPrefix) + Trim(Screen.ActiveForm.txtContractNo)
                End Select
                
                !ContractNo = !ContractNo + 1
                .Update
                .Requery

        End With
        
rsCONTROL.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub enableRECORD()
On Error GoTo err
        With frmODASMAllocation
            .txtAgreementDate.Locked = True
            .DTPickerAgreementDate.Enabled = True
            .txtContractNo.Locked = True
            .txtLandLordNo.Locked = True
            .txtNames.Locked = True
            .chkLeaseAll.Enabled = True
            .txtSignedBy.Locked = True
            .txtWitnessLandLord.Locked = False
            .txtWitnessCoy.Locked = True
            
        End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub updateRECORD()
       
    If ValidRecord = True Then
            saveRecord
            If k = 0 Then Exit Sub
            updateSITE
            updateLeasedPlotMasts
            upDateLeaseAnnualRent
            updateInstallments
            disableALLRECORD
            frmODASMAllocation.cmdPrint.Enabled = True
            frmODASMAllocation.Toolbar1.Buttons(3).Caption = "FINISH"
            frmODASMAllocation.Command1.Enabled = True
            frmODASMAllocation.Toolbar1.Buttons(2).Caption = "New &Record ": frmODASMAllocation.Toolbar1.Buttons(2).Image = 3
            getLANDLORDS
            showALLLandLORDSites
            showALLINSTALLMENTSDUE
    End If
End Sub

Public Sub SaveContractDetails()
    With frmODASMSiteRegistration
        Set rsFindRecord = New ADODB.Recordset
        If .txtContractNo.Text = "" Then GenerateContractNo
        
        rsFindRecord.Open "SELECT * FROM ODASMLeaseAgreement WHERE ContractNo='" & .txtContractNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF Or rsFindRecord.BOF Then
            rsFindRecord.AddNew
            rsFindRecord!ContractNo = .txtContractNo.Text
            rsFindRecord!dateprepared = Format(Date, "yyyy/MM/dd")
            rsFindRecord!Preparedby = CurrentUserName
        End If
        rsFindRecord!PlotNo = .txtPlotNo.Text
        rsFindRecord!AgreementDate = Format(.txtCommencementDate, "yyyy/MM/dd")
        rsFindRecord!AccountNo = .txtLandLordNo.Text
        rsFindRecord!CommencementDate = Format(.txtCommencementDate, "yyyy/MM/dd")
        rsFindRecord!expirydate = Format(.txtExpiryDate, "yyyy/MM/dd")
        rsFindRecord!AnnualRent = .txtAnnualRent
        rsFindRecord!PaymentMode = .cboPaymentMode
        rsFindRecord!LeaseDuration = .txtLeaseDuration
        rsFindRecord!AcquisitionDate = Format(.txtAcquisitionDate, "yyyy/MM/dd")
        rsFindRecord!AnnualRentIncrement = .txtAnnualRentIncrement
        rsFindRecord!RentVariationType = .txtRentVariationType
        rsFindRecord!AnnualRentIncrementType = .txtAnnualRentIncrementType
        rsFindRecord!IncrementStartYear = .txtIncrementStartYear
        rsFindRecord!IncrementFrequency = .txtIncrementFrequency
        rsFindRecord!WithLease = IIf(.chkWithLease.Value = vbChecked, "Y", "N")
        rsFindRecord!PaymentInterval = .txtPaymentInterval.Text
        rsFindRecord!Comments = .txtComments.Text
        
        rsFindRecord.Update
    End With
End Sub

Public Sub GenerateInstallmentPayable()
On Error GoTo err
    With frmODASMSiteRegistration
        
        Dim MaxInstallments, InstallmentAmount As Double
                    

        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" & .cboPaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF Or rsFindRecord.BOF Then
                MaxInstallments = 0
                installmentsAmount = 0
        Else
                MaxInstallments = CDbl(rsFindRecord!PaymentsInAYear) * CDbl(.txtLeaseDuration)
                installmentsAmount = CDbl(.txtAnnualRent.Text) / CDbl(rsFindRecord!PaymentsInAYear)
        End If
        
        For i = 1 To MaxInstallments
                
            Set rsSAVE = New ADODB.Recordset
            rsSAVE.Open "SELECT * FROM ODASMInstallment WHERE PlotNo='" & .txtPlotNo.Text & "' AND ContractNo = '" & .txtContractNo & "' and Installment = '" & i & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.EOF And rsSAVE.BOF Then
                    rsSAVE.AddNew
                    rsSAVE!PlotNo = .txtPlotNo
                    rsSAVE!ContractNo = .txtContractNo.Text
                    rsSAVE!Installment = i
                    rsSAVE!PaymentMode = .cboPaymentMode.Text
                    rsSAVE!InvoiceNo = Trim(.txtContractNo) + "-" + Trim(Str(i))
                    rsSAVE!AccountNo = .txtLandLordNo.Text
                    rsSAVE!PaymentFlag = "N"
                    rsSAVE!AmountPaid = 0
                    rsSAVE!Requisitioned = "N"
                    rsSAVE!Status = "INSTAL-CREATED"
                    rsSAVE!StatusDate = Date
                   
                   
            End If
            rsSAVE!PaymentDueDate = DateAdd("M", (i - 1) * rsFindRecord!CoverPeriod, .txtCommencementDate)
            rsSAVE!CurrentPeriod = getPeriod(rsSAVE!PaymentDueDate)
            rsSAVE!TotalRent = 0
            If ((i * rsFindRecord!CoverPeriod) Mod 12) <> 0 Then
                    rsSAVE!ContractYear = Fix((i * rsFindRecord!CoverPeriod) / 12) + 1
            Else
                    rsSAVE!ContractYear = Fix((i * rsFindRecord!CoverPeriod) / 12)
                    
                    '/ Increase the Rent Where appropriate
                    If rsSAVE!ContractYear >= CInt(.txtIncrementStartYear.Text) Then
                        If (rsSAVE!ContractYear) Mod .txtIncrementFrequency = 0 Then
                            If .txtAnnualRentIncrementType.Text = "P" Then
                                    rsSAVE!InstallmentPercent = installmentsAmount * (.txtAnnualRentIncrement / 100)
                            ElseIf .txtAnnualRentIncrementType.Text = "A" Then
                                    rsSAVE!InstallmentPercent = .txtAnnualRentIncrement
                            Else
                                    rsSAVE!InstallmentPercent = 0
                            End If
                            
                            If IsNull(rsSAVE!InstallmentPercent) Then
                                    rsSAVE!InstallmentPercent = 0
                                    installmentsAmount = CDbl(installmentsAmount)
                            Else: installmentsAmount = CDbl(installmentsAmount) + CDbl(rsSAVE!InstallmentPercent)
                            End If
                        End If
                    Else
                        rsSAVE!InstallmentPercent = 0
                    End If
            End If
        
            If (rsSAVE!ContractYear - 1) Mod .txtPaymentInterval = 0 Then
                  rsSAVE!PaymentDue = CDbl(installmentsAmount) * CDbl(.txtPaymentInterval)
                  rsSAVE!ContractLength = .txtPaymentInterval
            Else
                 rsSAVE!PaymentDue = 0
            End If
          
            If rsSAVE!AmountPaid > rsSAVE!PaymentDue Then
                    rsSAVE!AmountPaid = rsSAVE!PaymentDue
                     rsSAVE!PaymentDue = 0
                       rsSAVE!Balance = 0
            End If
              rsSAVE!PaymentDue = rsSAVE!PaymentDue - rsSAVE!AmountPaid
            If rsSAVE!PaymentDue = 0 Then
                rsSAVE!Requisitioned = "Y"
                If rsSAVE!AmountPaid > 0 Then
                    rsSAVE!PaymentFlag = "Y"
                End If
            End If
            
            
            rsSAVE!TotalRent = CDbl(rsSAVE!TotalRent) + CDbl(installmentsAmount)
         rsSAVE!PaymentDue = CDbl(installmentsAmount)
            rsSAVE!Balance = CDbl(rsSAVE!PaymentDue)
            Debug.Print installmentsAmount
            Debug.Print i
            rsSAVE.Update
        Next i
        
        Set rsSAVE = Nothing: Set rsFindRecord2 = Nothing: Set rsFindRecord = Nothing: Set rsFindRecord1 = Nothing

     End With
Exit Sub
err:
ErrorMessage
End Sub
Public Sub updateInstallments()
On Error GoTo err
    With frmODASMAllocation
        Set rsFindRecord2 = New ADODB.Recordset
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASMLeaseAgreement WHERE ContractNo = '" & .txtContractNo.Text & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
        
        If bAllowProcess = True Then
                rsFindRecord2.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" & .cboPaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        Else: rsFindRecord2.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" & rsFindRecord!PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        End If
        PaymentsInAYear = rsFindRecord2!PaymentsInAYear
Verify:
        Set rsFindRecord1 = New ADODB.Recordset
        rsFindRecord1.Open "SELECT * FROM ODASPInstallment WHERE ODASPInstallment.PaymentMode = '" & rsFindRecord!PaymentMode & "' and ODASPInstallment.LeasePeriod = '1';", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsFindRecord1.RecordCount = 0 Then
            
                Set rsFindRecord2 = New ADODB.Recordset
                rsFindRecord2.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" & rsFindRecord!PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                Dim PInts, Dur As String
                PInts = 1: Dur = 0
                
                Do While PInts <> (12 / rsFindRecord2!CoverPeriod) + 1
                    rsFindRecord1.AddNew
                    rsFindRecord1!LeasePeriod = 1
                    rsFindRecord1!PaymentMode = rsFindRecord!PaymentMode
                    rsFindRecord1!Installment = PInts
                    rsFindRecord1!Duration = Dur
                    rsFindRecord1!InstallmentDescription = "Installment" & " " & PInts
                    rsFindRecord1!dateprepared = Date
                    rsFindRecord1!Preparedby = CurrentUserName
                    
                    rsFindRecord1.Update
                    PInts = PInts + 1: Dur = Dur + rsFindRecord2!CoverPeriod
                Loop
                
                GoTo Verify
        
        End If
        
        Dim installmentsAmount As String, LeaseP, Rent As Variant
        LeaseP = 1: Rent = rsFindRecord!AnnualRent: ComDate = rsFindRecord!CommencementDate

        Do While LeaseP < rsFindRecord!LeaseDuration + 1
            installmentsAmount = Rent / PaymentsInAYear

            rsFindRecord1.MoveFirst
            
            While rsFindRecord1.EOF <> True
                    Set rsSAVE = New ADODB.Recordset
                    Dim InstallmentNo As String
                    
                    InstallmentNo = generateInstallmentNo
                    rsSAVE.Open "SELECT * FROM ODASMInstallment WHERE ContractYear = '" & LeaseP & "' and ContractNo = '" & rsFindRecord!ContractNo & "' and Installment = '" & rsFindRecord1!Installment & "' and PaymentMode = '" & rsFindRecord!PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    If rsSAVE.EOF And rsSAVE.BOF Then
                            rsSAVE.AddNew
                            rsSAVE!ContractNo = rsFindRecord4!ContractNo
                            rsSAVE!AccountNo = .txtLandLordNo.Text
                            rsSAVE!Installment = rsFindRecord1!Installment
                            rsSAVE!PaymentMode = .cboPaymentMode.Text
                            rsSAVE!InvoiceNo = InstallmentNo
                            rsSAVE!ContractYear = LeaseP
                            rsSAVE!TotalRent = Rent
                            rsSAVE!CurrentPeriod = CurrentPeriod
                            rsSAVE!PaymentDueDate = DateAdd("M", rsFindRecord1!Duration, ComDate)
                            rsSAVE!InstallmentPercent = rsFindRecord1!InstallmentPercent
                            rsSAVE!PaymentDue = CDbl(installmentsAmount)
                            rsSAVE!Balance = CDbl(installmentsAmount)
                            rsSAVE!PaymentFlag = "N"
                    End If
                    
                    rsSAVE.Update
                    rsFindRecord1.MoveNext
            Wend
            
            LeaseP = LeaseP + 1: ComDate = DateAdd("yyyy", 1, ComDate)
            If .optPercentage.Value = True Then
                Rent = Rent * (100 + rsFindRecord!PercentageIncreament) / 100
            ElseIf .optAmount.Value = True Then
                Rent = Rent + rsFindRecord!PercentageIncreament
            Else
                Rent = Rent * (100 + rsFindRecord!PercentageIncreament) / 100
            End If
            Loop
        Set rsSAVE = Nothing: Set rsFindRecord2 = Nothing: Set rsFindRecord = Nothing: Set rsFindRecord1 = Nothing

     End With
Exit Sub
err:
ErrorMessage
End Sub

Function generateInstallmentNo()
    With frmODASMAllocation
        Set rsFindRecord4 = New ADODB.Recordset
        rsFindRecord4.Open "SELECT * FROM ODASMLeaseagreement WHERE ContractNo = '" & .txtContractNo.Text & "' and AccountNo = '" & .txtLandLordNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord4.RecordCount = 0 Then Exit Function
        Set rsFindRecord2 = New ADODB.Recordset
        rsFindRecord2.Open "SELECT * FROM ODASMInstallment WHERE ContractNo = '" & rsFindRecord4!ContractNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord2.EOF And rsFindRecord2.BOF Then
            generateInstallmentNo = rsFindRecord4!ContractNo & "-" & 1
        Else
            generateInstallmentNo = rsFindRecord2!ContractNo & "-" & (rsFindRecord2.RecordCount + 1)
        End If
    End With
End Function
Public Sub upDateLeaseAnnualRent()
On Error GoTo err
    With frmODASMAllocation
        Set rsLease = New ADODB.Recordset
        rsLease.Open "SELECT sum(ODASPPlotMast.AnnualRent)as RentPayable, ODASPPlotMast.ExpiryDate,ODASPPlotMast.CommencementDate,ODASPPlotMast.ContractNo, ODASPPlotMast.LeaseDuration  FROM ODASPPlotMast WHERE ContractNo = '" & .txtContractNo.Text & "' Group By ContractNo, LeaseDuration,CommencementDate,ExpiryDate;", cnCOMMON, adOpenKeyset, adLockOptimistic
            
        Set rsLease1 = New ADODB.Recordset
        rsLease1.Open "SELECT * FROM ODASMLeaseAgreement WHERE ContractNo = '" & rsLease!ContractNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsLease1.RecordCount = 0 Then Exit Sub
            rsLease1!AnnualRent = rsLease!RentPayable
            rsLease1!LeaseDuration = rsLease!LeaseDuration
            rsLease1!CommencementDate = rsLease!CommencementDate
            rsLease1!expirydate = rsLease!expirydate
        rsLease1.Update
    End With
Exit Sub
err:
ErrorMessage
End Sub

Public Sub updateLeasedPlotMasts()
On Error GoTo err
    With frmODASMAllocation
    Dim search As String
        j = .ListView3.ListItems.Count
        For i = 1 To j
            If .ListView3.ListItems(i).Checked = True Then
            search = .ListView3.ListItems(i).Text
            
            Set rsFindRecord = New ADODB.Recordset
            rsFindRecord.Open "SELECT * FROM ODASPPlotMast WHERE MastNo = '" & search & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            If rsFindRecord.RecordCount = 0 Then Exit Sub
                rsFindRecord!LeasePrepared = "Y"
                rsFindRecord!ContractNo = .txtContractNo.Text
           
            rsFindRecord.Update
            End If
        Next i
    End With
Exit Sub
err:
ErrorMessage
End Sub

Public Sub updateSITE()
On Error GoTo err

       With frmODASMAllocation
            
                strSQL = "select * from ODASPPlot Where PlotNo = '" & .txtPlotNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                
                If .chkDeallocate.Value = 1 Then
                        rsSAVE!Status = "UN-ALLOCATED"
                    Else: rsSAVE!Status = "SITE-ACQUIRED"
                End If
                
                rsSAVE!AccountNo = .txtLandLordNo.Text
                rsSAVE.Update
                
                Set rsSAVE = Nothing
                strSQL = Empty
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Function ValidRecord()
On Error GoTo err
    ValidRecord = False
    With frmODASMAllocation
            If .txtPlotNo.Text <= "" Then
                strMessage = "The Site Number is required ..........."
                .txtPlotNo.SetFocus
            
            ElseIf .txtAgreementDate.Text = Empty Then
                strMessage = "The Date Assigned cannot not be ignored"
                .txtAgreementDate.SetFocus
            
            ElseIf .txtLandLordNo.Text = Empty Then
                strMessage = "The Land Lord is Required ..........."
                .txtLandLordNo.SetFocus
            
            ElseIf .chkYes.Value = 1 And .optPercentage.Value = False And .optAmount.Value = False Then
                strMessage = "Specify the whether annual rent grows by a % or amount..."
            
            ElseIf .txtpaymentMode.Text = Empty Then
                strMessage = "The Payment Mode is Required ..........."
                .txtpaymentMode.SetFocus
            
            ElseIf .chkYes.Value = 1 And .txtPercentage.Text = Empty Then
                strMessage = "The rent increament value is Required ...."
                .txtPercentage.SetFocus
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


