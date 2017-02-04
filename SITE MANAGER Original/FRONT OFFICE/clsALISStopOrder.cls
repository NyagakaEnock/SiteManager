VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALISStopOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rsLOADGRID As clsALISGRID
Dim rsREFERENCE As ADODB.Recordset, numRECORDS As Integer




Private Sub updateDepartment()
'On Error GoTo err

        Dim rsDEPARTMENT As ADODB.Recordset, strDEPARTMENT As String
        Set rsDEPARTMENT = New Recordset
        
        strDEPARTMENT = "SELECT * FROM ALISPDepartment WHERE AccountNo = '" & (frmALISMDeductionForm.cboAccountNo.Text) & "' and DepartmentCode = '" & frmALISMDeductionForm.cboDepartmentCode.Text & "';"
        rsDEPARTMENT.Open strDEPARTMENT, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsDEPARTMENT
                If .EOF And .BOF Then Exit Sub
                !DiskettePrepared = "Y"
                !NewBusinessCreated = "Y"
                !RenewalCreated = "Y"
                If bDeleteRECORd = True Then
                        !MonthPrepared = ""
                Else: !MonthPrepared = frmALISMDeductionForm.txtMonth.Text
                End If
                !dateprepared = Date
                !Preparedby = CurrentUserName
                .Update
                .Requery
        End With

rsDEPARTMENT.Close
strDEPARTMENT = ""

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub RedoDepartment()
'On Error GoTo err

        Dim rsDEPARTMENT As ADODB.Recordset, strDEPARTMENT As String
        Set rsDEPARTMENT = New Recordset
        
        strDEPARTMENT = "SELECT * FROM ALISPDepartment WHERE AccountNo = '" & (frmALISMDeductionForm.cboAccountNo.Text) & "' and DepartmentCode = '" & frmALISMDeductionForm.cboDepartmentCode.Text & "';"
        rsDEPARTMENT.Open strDEPARTMENT, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsDEPARTMENT
                If .EOF And .BOF Then Exit Sub
                !MonthPrepared = "00"
                frmALISMDeductionForm.txtMonth.Text = !MonthPrepared
                .Update
                .Requery
        End With

rsDEPARTMENT.Close
strDEPARTMENT = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdAddNew_Click()
    clearALLRECORD
    enableALLRECORD
    disableButtons
    loadDEFAULTS
End Sub


Private Sub cmdCancel_Click()
    enableButtons
    disableALLRECORD
    clearALLRECORD
End Sub

Private Sub cmdDelete_Click()
    With frmALISMDeductionForm
        If .txtMonth.Text <= "" Or .cboAccountNo <= "" Or .cboDepartmentCode <= "" Then
                MsgBox "Kindly Load the Record"
        Else: updateDepartment
                clearALLRECORD
        End If

    End With
End Sub

Private Sub cmdEdit_Click()
    If frmALISMDeductionForm.cboDepartmentCode.Text >= "" Then
            RedoDepartment
    End If
End Sub

Private Sub cmdPendingDepts_Click()
    If Screen.ActiveForm.cboAccountNo.Text <= "" Then
        MsgBox "The Employer Code is Mandatory", vbOKOnly
        Screen.ActiveForm.cboAccountNo.SetFocus
    ElseIf Screen.ActiveForm.txtMonth.Text <= "" Then
        MsgBox "The Month cannot be left Blank", vbOKOnly
        Screen.ActiveForm.txtMonth.SetFocus
    Else
            Load frmALISRPendingDepts
            frmALISRPendingDepts.Show 1, Me
    End If
End Sub

Private Sub cmdUpdate_Click()
    updateRECORD
    If bsaveRECORD = True Then
            deleteTSCDATA
            GenerateForm
            If bsaveRECORD = False Then
                updateDepartment
                disableALLRECORD
                enableButtons
            End If
    End If
    
    Set rsLOADGRID = New clsALISGRID
    rsLOADGRID.loadDEPARTMENTGRID
    rsLOADGRID.loadPENDINGDEPTGRID
    rsLOADGRID.loadPREPAREDDEPTGRID
    Set rsLOADGRID = Nothing

    
End Sub
Private Sub deleteTSCDATA()
'On Error GoTo err

        Dim strTSC As String, rsTSC, rsDEL As ADODB.Recordset, multiply As Currency
        Set rsTSC = New Recordset

        strTSC = "select * from ALISMTSCData; "
        rsTSC.Open strTSC, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsTSC
            If .RecordCount > 0 Then
                    Set rsDEL = New ADODB.Recordset
                    rsDEL.Open "Delete from ALISMTSCData;", cnCOMMON, adOpenKeyset, adLockOptimistic
            End If
        End With

rsTSC.Close
strTSC = ""
        
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub updateRECORD()
'On Error GoTo err
    frmALISMDeductionForm.MousePointer = vbHourglass
    
    If Screen.ActiveForm.cboDepartmentCode.Text <= "" Then
        MsgBox "The Department Code cannot be Left Blank", vbOKOnly
        Screen.ActiveForm.cboDepartmentCode.SetFocus
    
    ElseIf Screen.ActiveForm.cboAccountNo.Text <= "" Then
        MsgBox "The Employer Code is Mandatory", vbOKOnly
        Screen.ActiveForm.cboAccountNo.SetFocus
    
    ElseIf Screen.ActiveForm.txtMonth.Text <= "" Then
        MsgBox "The Month cannot be left Blank", vbOKOnly
        Screen.ActiveForm.txtMonth.SetFocus
    Else
            bsaveRECORD = True
    End If
    
    
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdPrint_Click()
'On Error GoTo err

    If Screen.ActiveForm.cboDepartmentCode.Text <= "" Then
        MsgBox "The Department Code cannot be Left Blank", vbOKOnly
        Screen.ActiveForm.cboDepartmentCode.SetFocus
    ElseIf Screen.ActiveForm.cboAccountNo.Text <= "" Then
        MsgBox "The Employer Code is Mandatory", vbOKOnly
        Screen.ActiveForm.cboAccountNo.SetFocus
    ElseIf Screen.ActiveForm.txtMonth.Text <= "" Then
        MsgBox "The Month cannot be left Blank", vbOKOnly
        Screen.ActiveForm.txtMonth.SetFocus
    Else
    
        Dim rsSTOPORDER As ADODB.Recordset, strSTOPORDER As String
        Set rsSTOPORDER = New ADODB.Recordset
                            
        strSTOPORDER = "SELECT * FROM ALISMDeduction where monthCode = '" & frmALISMDeductionForm.txtMonth.Text & "' and AccountNo = '" & frmALISMDeductionForm.cboAccountNo.Text & "' and departmentCode = '" & frmALISMDeductionForm.cboDepartmentCode.Text & "' ;"
        rsSTOPORDER.Open strSTOPORDER, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        Set rptALISStopOrder.DataSource = rsSTOPORDER
        rptALISStopOrder.Show
    End If

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdSearch_Click()
'On Error GoTo Myerr
        Dim strQRE, strQRE1, strQRE2 As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
            strQRE = InputBox("Enter Employer Code Code to search.", "Search Value")
            strQRE1 = InputBox("Enter The Department Coode to Search", "Search Value")
            strQRE2 = InputBox("Enter The Month to Print", "Search Value")
            
            rsFind.Open "SELECT * FROM ALISMDeduction, ALISPdepartment WHERE ALISMDeduction.AccountNo = '" & strQRE & "' and ALISMdeduction.DepartmentCode = '" & strQRE1 & "' and ALISMDeduction.monthcode = '" & strQRE2 & "' and ALISMdeduction.departmentcode = ALISPdepartment.departmentCode ;", cnCOMMON, adOpenKeyset, adLockOptimistic

            With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else:
                        
                            frmALISMDeductionForm.cboDepartmentCode.Text = !DepartmentCode & ""
                            frmALISMDeductionForm.cboAccountNo.Text = !AccountCode
                            frmALISMDeductionForm.txtMonth.Text = !MonthCode
                            frmALISMDeductionForm.txtAmountOff.Text = !AmountOff
                            frmALISMDeductionForm.txtAmountOn.Text = !Amounton
                            frmALISMDeductionForm.txtBalance.Text = !Balance
                            frmALISMDeductionForm.txtCheckDigit.Text = !CheckDigit
                            frmALISMDeductionForm.txtDeduction.Text = !Deduction
                            frmALISMDeductionForm.txtDeductionCode.Text = !DeductionCode
                            frmALISMDeductionForm.txtCode.Text = !code
                            frmALISMDeductionForm.txtPersonalNumber.Text = !PersonalNumber
                            frmALISMDeductionForm.txtRecurrentAmount.Text = !NonRecurrentAmount
                            Edit = True
                End If
 
            End With

Exit Sub

Myerr:
                    ErrorMessage


End Sub





Private Sub DTPickerReferenceDate_Change()
'On Error GoTo err
    With frmALISMDeductionForm
        .txtReferenceDate.Text = .DTPickerReferenceDate.Value
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
        frmALISMDeductionForm.SSTabDeduction.Tab = 0
    
        clearALLRECORD
        enableButtons
        disableALLRECORD
        
End Sub

Private Sub Form_Load()
        OpenConnection
End Sub

Private Sub loadDEFAULTS()
'On Error GoTo err
        frmALISMDeductionForm.DTPickerReferenceDate.Value = "31 / 12 / 2001"
        frmALISMDeductionForm.txtReferenceDate.Text = frmALISMDeductionForm.DTPickerReferenceDate.Value

        frmALISMDeductionForm.txtMonth.Text = Month(Date) + 1
        
        If Len(frmALISMDeductionForm.txtMonth.Text) = 1 Then
            frmALISMDeductionForm.txtMonth.Text = "0" + frmALISMDeductionForm.txtMonth.Text
        End If
        frmALISMDeductionForm.ProgressBar2.Visible = False
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub txtMonth_LostFocus()
'On Error GoTo err

        If Screen.ActiveForm.txtMonth.Text > "12" Then
            MsgBox "The Month Cannot be Greater Than 12", vbOKOnly
            Screen.ActiveForm.txtMonth.SetFocus
        ElseIf Screen.ActiveForm.txtMonth.Text <= "00" Then
            MsgBox "The Month MUST be Greater that 00 ", vbOKOnly
            Screen.ActiveForm.txtMonth.SetFocus
        Else
                Set rsLOADGRID = New clsALISGRID
                rsLOADGRID.loadPENDINGDEPTGRID
                Set rsLOADGRID = Nothing

        End If
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub GenerateForm()
'On Error GoTo err
        
        '/ Procedure to used to generate policy deduction Form

        strSQL = ""
        numRECORDS = 0
        Set rsREFERENCE = New ADODB.Recordset

        strSQL = "SELECT * FROM ODASMJobBrief, ODASPAccount, ODASPAccount where ODASPAccount.AccountNo = '" & Screen.ActiveForm.cboAccountNo.Text & "' and ODASPAccount.DepartmentCode = '" & Screen.ActiveForm.cboDepartmentCode.Text & "' and ODASMJobBrief.AccountNo LIKE ODASPAccount.AccountNo and ODASPAccount.AccountNo = ODASPAccount.AccountNo and ODASMJobBrief.StatusCode = 'INFORCE';"
        rsREFERENCE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        With rsREFERENCE
                If .BOF Or .EOF Then Exit Sub
                GenerateDeductionForm
        End With
        
                
        rsREFERENCE.Close
        strSQL = ""
        
        '/ Procedure is used to generate proposal deductions
        
        Set rsREFERENCE = New ADODB.Recordset
        strSQL = "SELECT * FROM ALISMProposal, ODASPAccount, ODASPAccount where ODASPAccount.AccountNo = '" & Screen.ActiveForm.cboAccountNo.Text & "' and ODASPAccount.DepartmentCode = '" & Screen.ActiveForm.cboDepartmentCode.Text & "' and ALISmProposal.AccountNo LIKE ODASPAccount.AccountNo and ODASPAccount.AccountNo = ODASPAccount.AccountNo and (ALISMProposal.statusCode = 'COMPLETE' );"
        
        rsREFERENCE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic


        With rsREFERENCE
                If .BOF Or .EOF Then Exit Sub
                GenerateDeductionForm
        End With
        
        rsREFERENCE.Close
        strSQL = ""
        
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub GenerateDeductionForm()
'On Error GoTo err

        With rsREFERENCE
            If .BOF Or .EOF Then Exit Sub
            
            numRECORDS = .RecordCount
            
            With frmALISMDeductionForm.ProgressBar2
                        .Visible = True
                        .Max = numRECORDS
                        .Min = 0
                        .Value = 0
            End With
            
            Do Until .EOF
                    If IsNull(!AccountNo) = True Then GoTo test:
                    If DateDiff("D", !MaturityDate, frmALISMDeductionForm.txtReferenceDate.Text) > 0 Then GoTo test:
                    'If !UnitCount >= (!TermOfPolicy * 12) Then GoTo test:

                    frmALISMDeductionForm.txtCheckDigit.Text = CDbl(!AccountNo) Mod 10
                    frmALISMDeductionForm.txtPersonalNumber.Text = Trim(!AccountNo)
                    
                    Select Case Len(!AccountNo)
                        Case 3:  frmALISMDeductionForm.txtPersonalNumber.Text = "0000000" + Trim(!AccountNo)
                        Case 4:  frmALISMDeductionForm.txtPersonalNumber.Text = "000000" + Trim(!AccountNo)
                        Case 5:  frmALISMDeductionForm.txtPersonalNumber.Text = "00000" + Trim(!AccountNo)
                        Case 6:  frmALISMDeductionForm.txtPersonalNumber.Text = "0000" + Trim(!AccountNo)
                        Case 7:  frmALISMDeductionForm.txtPersonalNumber.Text = "000" + Trim(!AccountNo)
                        Case 8:  frmALISMDeductionForm.txtPersonalNumber.Text = "00" + Trim(!AccountNo)
                        Case 9:  frmALISMDeductionForm.txtPersonalNumber.Text = "0" + Trim(!AccountNo)
                        Case 8:  frmALISMDeductionForm.txtPersonalNumber.Text = Trim(!AccountNo)
                    End Select
                    
                    frmALISMDeductionForm.txtDeductionCode.Text = Trim(!DeductionCode)
                    frmALISMDeductionForm.txtItemAccountCode.Text = Trim(!ItemAccountCode)
                    
                    frmALISMDeductionForm.txtAmountOn.Text = Str(Fix(!ExpectedPremium)) + "00"

                    Select Case Len(Trim(frmALISMDeductionForm.txtAmountOn.Text))
                        Case 4: frmALISMDeductionForm.txtAmountOn.Text = "0000" + Trim(frmALISMDeductionForm.txtAmountOn.Text)
                        Case 5: frmALISMDeductionForm.txtAmountOn.Text = "000" + Trim(frmALISMDeductionForm.txtAmountOn.Text)
                        Case 6: frmALISMDeductionForm.txtAmountOn.Text = "00" + Trim(frmALISMDeductionForm.txtAmountOn.Text)
                        Case 7: frmALISMDeductionForm.txtAmountOn.Text = "0" + Trim(frmALISMDeductionForm.txtAmountOn.Text)
                        Case 8: frmALISMDeductionForm.txtAmountOn.Text = Trim(frmALISMDeductionForm.txtAmountOn.Text)
                    End Select
                    
                    frmALISMDeductionForm.txtCode.Text = " "
                    frmALISMDeductionForm.txtRecurrentAmount.Text = "        "
                    frmALISMDeductionForm.txtAmountOff.Text = "        "
                    frmALISMDeductionForm.txtReference.Text = (Right(Trim(!ProposalNo), 6))
                    
                    Select Case Len(Trim(frmALISMDeductionForm.txtReference.Text))
                        Case 1: frmALISMDeductionForm.txtReference.Text = "0000000" + Trim(frmALISMDeductionForm.txtReference.Text)
                        Case 2: frmALISMDeductionForm.txtReference.Text = "000000" + Trim(frmALISMDeductionForm.txtReference.Text)
                        Case 3: frmALISMDeductionForm.txtReference.Text = "00000" + Trim(frmALISMDeductionForm.txtReference.Text)
                        Case 4: frmALISMDeductionForm.txtReference.Text = "0000" + Trim(frmALISMDeductionForm.txtReference.Text)
                        Case 5: frmALISMDeductionForm.txtReference.Text = "000" + Trim(frmALISMDeductionForm.txtReference.Text)
                        Case 6: frmALISMDeductionForm.txtReference.Text = "00" + Trim(frmALISMDeductionForm.txtReference.Text)
                        Case 7: frmALISMDeductionForm.txtReference.Text = "0" + Trim(frmALISMDeductionForm.txtReference.Text)
                        Case 8: frmALISMDeductionForm.txtReference.Text = Trim(frmALISMDeductionForm.txtReference.Text)
                    End Select

                    
                    frmALISMDeductionForm.txtBalance.Text = "            "
                    frmALISMDeductionForm.txtIdentityDetails.Text = "                                                          "
                    
                    With frmALISMDeductionForm
                        .txtDeduction.Text = .txtMonth + Trim(.cboDepartmentCode) + .txtCheckDigit + .txtPersonalNumber + .txtDeductionCode + .txtItemAccountCode + .txtAmountOn + .txtCode + .txtRecurrentAmount + .txtAmountOff + .txtReference + .txtBalance + .txtIdentityDetails.Text
                    End With
                                        
                   'This process is used to update the Deduction file with the relevant Data
                    
                    Dim rsDEDUCTION As ADODB.Recordset, strDEDUCTION As String
                    Set rsDEDUCTION = New Recordset
                    
                    strDEDUCTION = "SELECT * FROM ALISMDeduction where MonthCode = '" & frmALISMDeductionForm.txtMonth.Text & "' and ALISMdeduction.AccountNo = '" & frmALISMDeductionForm.txtReference.Text & "'; "
                    
                    rsDEDUCTION.Open strDEDUCTION, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    With rsDEDUCTION
                        
                        If .EOF Or .BOF Then
                            .AddNew
                            !MonthCode = frmALISMDeductionForm.txtMonth.Text
                            !AccountNo = frmALISMDeductionForm.txtReference.Text
                        End If
                        
                        !Amounton = CDbl(frmALISMDeductionForm.txtAmountOn.Text)
                        !AmountOff = CDbl(frmALISMDeductionForm.txtAmountOff.Text)
                        !PersonalNumber = frmALISMDeductionForm.txtPersonalNumber.Text
                        !DeductionCode = frmALISMDeductionForm.txtDeductionCode.Text
                        !Deduction = frmALISMDeductionForm.txtDeduction.Text
                        !DepartmentCode = Trim(frmALISMDeductionForm.cboDepartmentCode.Text)
                        !AccountCode = frmALISMDeductionForm.cboAccountNo
                        !Balance = frmALISMDeductionForm.txtBalance.Text
                        !CheckDigit = frmALISMDeductionForm.txtCheckDigit.Text
                        !code = frmALISMDeductionForm.txtCode.Text
                        !IdentityDetails = frmALISMDeductionForm.txtIdentityDetails.Text
                        !NonRecurrentAmount = frmALISMDeductionForm.txtRecurrentAmount.Text
                        !ItemAccountCode = frmALISMDeductionForm.txtItemAccountCode.Text
                        !dateprepared = Date
                        !Preparedby = CurrentUserName
                        .Update
                    End With
                    
                    
                    Dim rsTSC As ADODB.Recordset, strTSC As String
                    Set rsTSC = New Recordset
                    
                    strTSC = "SELECT * FROM ALISMTSCData where Deduction = '" & frmALISMDeductionForm.txtDeduction.Text & "'; "
                    rsTSC.Open strTSC, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    With rsTSC
                        
                        If .EOF Or .BOF Then
                            .AddNew
                            !Deduction = frmALISMDeductionForm.txtDeduction.Text
                            .Update
                        End If
                    End With
test:
                    .MoveNext
                    frmALISMDeductionForm.ProgressBar2.Value = frmALISMDeductionForm.ProgressBar2.Value + 1
            Loop
            frmALISMDeductionForm.ProgressBar2.Visible = False
            bsaveRECORD = False
       End With


     frmALISMDeductionForm.MousePointer = 0

Exit Sub

err:
    ErrorMessage
End Sub

