VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalculateCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public rsPOLICY, rsPAYMENTMODE, rsAGENTNO, rsSUPERVISOR, rsPREMPAY As ADODB.Recordset
Public strPOLICY, strPAYMENTMODE, strAGENTNO, strSUPERVISOR, strPREMPAY As String
Dim rsBATCHPOL As ADODB.Recordset
Dim MYCOUNT As Integer
Public Sub validateCOMMISSION()
'On Error GoTo err
        With Screen.ActiveForm
            bsaveRECORD = False
            
            If .cboDocumentNo.Text = Empty Then
                    MsgBox "The Document number Cannot be Left Blank"
                    .cboDocumentNo.SetFocus
            ElseIf CDbl(.txtCommissionRecords) > 0 Then
                    MsgBox "The Commission Has Already been calculated"
                    .txtCommissionRecords.SetFocus
            ElseIf CDbl(.txtExpectedAmount.Text) <= 0 Then
                    MsgBox "The expected Premium MUST be Greater Than Zero"
                    .txtExpectedAmount.SetFocus
            ElseIf CDbl(.txtunitsPaid.Text) <= 0 Then
                    MsgBox "The Units Paid MUST be Greater Than Zero"
                    .txtunitsPaid.SetFocus
            ElseIf CDbl(.txtUnitCount.Text) <= 0 Then
                    MsgBox "The unit Count MUST be Less than Zero"
                    .txtUnitCount.SetFocus
            ElseIf .txtAccountingPeriod.Text = Empty Then
                    MsgBox "The Period Cannot be Empty"
                    .txtAccountingPeriod.SetFocus
            ElseIf CDbl(.txtReceiptAmount) < CDbl(.txtExpectedAmount) Then
                    MsgBox "The Receipt Amount is Less than the Expected Amount, CANNOT calculate Commission"
                    .txtExpectedAmount.SetFocus
            ElseIf .txtAgentStatus.Text = "T" Then
                    MsgBox "The Agent Status Has been terminated"
                    .txtAgentStatus.SetFocus
            Else
                    bsaveRECORD = True
            End If
            
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub CalculateCommission()
'On Error GoTo err
        
        '/ obtain the necessary Policies details required for the processing
        '/ of commission
        
        Set rsPOLICY = New ADODB.Recordset
                                
        strPOLICY = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo  = '" & Screen.ActiveForm.cboDocumentNo & " ' and (statusCode = 'INFORCE' or statuscode = 'UNISSUED');"
        rsPOLICY.Open strPOLICY, cnCOMMON, adOpenKeyset, adLockOptimistic

        '/* This record set is used to obtain the cover Period for purcpose
        '/* of calculating the commission especially for non-monthly payments
        
        If rsPOLICY.EOF Or rsPOLICY.BOF = True Then Exit Sub
        
        Set rsPAYMENTMODE = New ADODB.Recordset
            
        rsPAYMENTMODE.Open "SELECT * FROM ALISPPaymentMode where PaymentMode = '" & rsPOLICY!PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsPAYMENTMODE.EOF Or rsPAYMENTMODE.BOF = True Then
                MsgBox "The Payment Mode has not been set, The Program is exiting ", vbOKOnly
                Exit Sub
        End If
        
        '/ Used to obtain the Agent who brought the business to the insuarance
        '/ company
        
        Set rsAGENTNO = New ADODB.Recordset
                                        
        strAGENTNO = "SELECT * FROM ALISPAgent where AgentNo = '" & rsPOLICY!Agentno & "';"
        rsAGENTNO.Open strAGENTNO, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsAGENTNO.EOF Or rsAGENTNO.BOF = True Then
                MsgBox "The Agent Does not Exist within the Database, might have been accidentally Deleted", vbOKOnly
                Exit Sub
        End If
        
        '/ Used to obtain the Agent's Supervisior who brought the business to the insuarance
        '/ company
        
        If rsPOLICY!supervisorNo > "" Then
        
            Set rsSUPERVISOR = New ADODB.Recordset
                                            
            strSUPERVISOR = "SELECT * FROM ALISMAgentsPay where AgentNo = '" & rsPOLICY!supervisorNo & "';"
            rsSUPERVISOR.Open strSUPERVISOR, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSUPERVISOR.EOF Or rsSUPERVISOR.BOF = True Then
                    MsgBox "The Supervisor Agent Does not Exist within the Database, might have been accidentally Deleted", vbOKOnly
                    'Exit Sub
            End If

        End If
                '/* end if the Supervisor Number has not been entered */
        '/* obtain all parameters that relate to calculation of agents
        '/* Payments based on the premium
        
        Set rsPREMPAY = New ADODB.Recordset
            
        strPREMPAY = "SELECT * FROM ALISPAgentsPay WHERE GeneratedFrom  = 'P' and ProductCode = '" & rsPOLICY!ProductCode & "';"
        rsPREMPAY.Open strPREMPAY, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsPREMPAY
                If .EOF Or .BOF = True Then Exit Sub
                
                '/ This procedure is used to calculate the payment to an Agent that are generated from the premium. Such payment include
                '/ commission and persistency bonus that are paid based on the premium.
                
                Do Until .EOF
                                      
                     If Val(Screen.ActiveForm.txtUnitCount) >= !MinimumPremiums And Val(Screen.ActiveForm.txtUnitCount) <= !MaximumPremiums Then
                                                       
                            '/ This procedure is used to calculate the commission for the agent
                            '/ based on the premium received
                                
                            saveCOMMISSION
                                                                    
                            '/ The procedure is used to calculate the Payments to the Agents that are generateed
                            '/ from the commission calculated. This include such payment as  overide commission that are
                            '/ paid to the Agents Supervisior
                                    
                                                         
                            Dim rsCOMMISSION As ADODB.Recordset, strCommission As String
                            Set rsCOMMISSION = New Recordset
                                                
                            strCommission = "SELECT * FROM ALISPAgentsPay WHERE GeneratedFrom  = 'C' and ProductCode = '" & rsPOLICY!ProductCode & "' and minimumpremiums <= '" & Screen.ActiveForm.txtUnitCount & "' And MaximumPremiums <= ' " & Screen.ActiveForm.txtUnitCount & "' ;"
                            rsCOMMISSION.Open strCommission, cnCOMMON, adOpenKeyset, adLockOptimistic

                            With rsCOMMISSION
                                    
                                    If .EOF Or .BOF = True Then Exit Sub
                                          
                                        '/This procedure is used to calculate the commission for the agent
                                        '/ based on the premium received
                                        
                                        If rsPOLICY!supervisorNo > 0 Then
                                                saveCOMMISSION
                                        End If
                                        
                                        '/* End with rsPAY
                                            '/* if the payment to the Agent is made from the commission generated
                                        
                            End With
                                            '/* End with Commission
                     
                    '/* Calculate payment to agents that are not based on the premium count unlike the case
                    '/ Above. This will include such payment as premium tax that is a percentage
                    '/ of the premium paid
                    
                    ElseIf !MinimumPremiums = 0 And !MaximumPremiums = 0 Then
                        
                        saveCOMMISSION
                                                                                     
                        Dim rsCOMM As ADODB.Recordset, strComm As String
                        Set rsCOMM = New Recordset
                                            
                        strComm = "SELECT * FROM ALISPAgentsPay WHERE GeneratedFrom  = 'C' and ProductCode = '" & rsPOLICY!ProductCode & "';"
                        rsCOMM.Open strComm, cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsCOMM
                            'If Val(screen.activeform.txtUnitCount) >= !MinimumPremiums And Val(screen.activeform.txtUnitCount) <= !MaximumPremiums Then
                                  
                                '/This procedure is used to calculate the commission for the agent
                                '/ based on the premium received
                                                  
                                    saveCOMMISSION
                                            
                                '/* End with rsPAY
                                '/* if the payment to the Agent is made from the commission generated
                                
                        End With
                                            '/* End with Commission

                    End If
                                        '/* End if The Payment Type to the Agent is based on the Premium Count
                    .MoveNext
                
                Loop
                
                                        '/* end if calculation based on Premium
        End With


Exit Sub

err:
    ErrorMessage
    
End Sub
Private Sub checkPAYMENTMODE()
'On Error GoTo err

            Set rsPAYMENTMODE = New ADODB.Recordset
                
            rsPAYMENTMODE.Open "SELECT * FROM ALISPPaymentMode where PaymentMode = '" & rsBATCHPOL!PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            With rsPAYMENTMODE
                    If .EOF Or .BOF = True Then
                            bCONTINUE = False
                            MsgBox "The Payment Mode has not been set, The Program is exiting ", vbOKOnly
                            Exit Sub
                    End If
            End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub checkAGENT()
'On Error GoTo err
                '/ Used to obtain the Agent who brought the business to the insuarance
                '/ company
                
                Set rsAGENTNO = New ADODB.Recordset
                                                
                strAGENTNO = "SELECT * FROM ALISPAgent where AgentNo = '" & rsBATCHPOL!Agentno & "';"
                rsAGENTNO.Open strAGENTNO, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsAGENTNO
                        If .EOF Or .BOF = True Then
                                bCONTINUE = False
                                MsgBox "The Agent Does not Exist within the Database, might have been accidentally Deleted", vbOKOnly
                                Exit Sub
                        End If
                End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub checkAGENTSUP()
'On Error GoTo err

        '/ Used to obtain the Agent's Supervisior who brought the business to the insuarance
        '/ company
        
        If rsBATCHPOL!supervisorNo > "" Then
        
            Set rsSUPERVISOR = New ADODB.Recordset
                                            
            strSUPERVISOR = "SELECT * FROM ALISMAgentsPay where AgentNo = '" & rsBATCHPOL!supervisorNo & "';"
            rsSUPERVISOR.Open strSUPERVISOR, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            With rsSUPERVISOR
                    If .EOF Or .BOF = True Then
                            'bCONTINUE = False
                            'MsgBox "The Supervisor Agent Does not Exist within the Database, might have been accidentally Deleted", vbOKOnly
                            Exit Sub
                    End If
            End With

        End If


Exit Sub
err:
    ErrorMessage
End Sub
Public Sub CalculateBatchCommission()
'On Error GoTo err
        
        
        '/ obtain the necessary Policies details required for the processing
        '/ of commission
        
        Set rsBATCHPOL = New ADODB.Recordset
        strSQL = ""
        strSQL = "SELECT * FROM ALISMReceiptDetails, ODASMJobBrief WHERE  ALISMReceiptDetails.DocumentNo = ODASMJobBrief.JobBriefNo and (ODASMJobBrief.statusCode = 'INFORCE' or ODASMJobBrief.statuscode = 'UNISSUED') and ALISMReceiptDetails.AccountingPeriod = '" & Screen.ActiveForm.cboCurrentPeriod.Text & "';"
        rsBATCHPOL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        '/* This record set is used to obtain the cover Period for purcpose
        '/* of calculating the commission especially for non-monthly payments
              
        With rsBATCHPOL
                If .EOF Or .BOF = True Then Exit Sub
                MYCOUNT = .RecordCount
                
                With frmALISRCommissionPeriod.ProgressBar1
                        .Visible = True
                        .Max = MYCOUNT
                        .Min = 0
                        .Value = 0
                End With
                
                Do While Not .EOF
                    bCONTINUE = True
                    checkPAYMENTMODE
                    checkAGENT
                    checkAGENTSUP
                    
                    If bCONTINUE = True Then
                            computeBATCHCOMMISSION
                    End If
                    
                    frmALISRCommissionPeriod.ProgressBar1.Value = frmALISRCommissionPeriod.ProgressBar1.Value + 1
                    .MoveNext
                Loop
        
        
                frmALISRCommissionPeriod.ProgressBar1.Visible = False
        End With
Exit Sub

err:
    ErrorMessage
    
End Sub
Private Sub computeBATCHCOMMISSION()
'On Error GoTo err
    
    '/* obtain all parameters that relate to calculation of agents
    '/* Payments based on the premium
    
    Set rsPREMPAY = New ADODB.Recordset
        
    strPREMPAY = "SELECT * FROM ALISPAgentsPay WHERE GeneratedFrom  = 'P' and ProductCode = '" & rsBATCHPOL!ProductCode & "';"
    rsPREMPAY.Open strPREMPAY, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsPREMPAY
            If .EOF Or .BOF = True Then Exit Sub
            
            '/ This procedure is used to calculate the payment to an Agent that are generated from the premium. Such payment include
            '/ commission and persistency bonus that are paid based on the premium.
            
            Do Until .EOF
                                  
                 If Val(rsBATCHPOL!UnitCount) >= !MinimumPremiums And Val(rsBATCHPOL!UnitCount) <= !MaximumPremiums Then
                                                   
                        '/ This procedure is used to calculate the commission for the agent
                        '/ based on the premium received
                            
                        saveBATCHCOMMISSION
                                                                
                        '/ The procedure is used to calculate the Payments to the Agents that are generateed
                        '/ from the commission calculated. This include such payment as  overide commission that are
                        '/ paid to the Agents Supervisior
                                
                                                     
                        Dim rsCOMMISSION As ADODB.Recordset, strCommission As String
                        Set rsCOMMISSION = New Recordset
                                            
                        strCommission = "SELECT * FROM ALISPAgentsPay WHERE GeneratedFrom  = 'C' and ProductCode = '" & rsBATCHPOL!ProductCode & "' and minimumpremiums <= '" & rsBATCHPOL!UnitCount & "' And MaximumPremiums <= ' " & rsBATCHPOL!UnitCount & "' ;"
                        rsCOMMISSION.Open strCommission, cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsCOMMISSION
                                
                                If .EOF Or .BOF = True Then Exit Sub
                                      
                                    '/This procedure is used to calculate the commission for the agent
                                    '/ based on the premium received
                                    
                                    If rsBATCHPOL!supervisorNo > 0 Then
                                            saveBATCHCOMMISSION
                                    End If
                                    
                                    '/* End with rsPAY
                                        '/* if the payment to the Agent is made from the commission generated
                                    
                        End With
                                        '/* End with Commission
                 
                '/* Calculate payment to agents that are not based on the premium count unlike the case
                '/ Above. This will include such payment as premium tax that is a percentage
                '/ of the premium paid
                
                ElseIf !MinimumPremiums = 0 And !MaximumPremiums = 0 Then
                    
                    saveBATCHCOMMISSION
                                                                                 
                    Dim rsCOMM As ADODB.Recordset, strComm As String
                    Set rsCOMM = New Recordset
                                        
                    strComm = "SELECT * FROM ALISPAgentsPay WHERE GeneratedFrom  = 'C' and ProductCode = '" & rsBATCHPOL!ProductCode & "';"
                    rsCOMM.Open strComm, cnCOMMON, adOpenKeyset, adLockOptimistic

                    With rsCOMM
                        'If Val(rsBATCHPOL!UnitCount) >= !MinimumPremiums And Val(rsBATCHPOL!UnitCount) <= !MaximumPremiums Then
                              
                            '/This procedure is used to calculate the commission for the agent
                            '/ based on the premium received
                                              
                                saveBATCHCOMMISSION
                                        
                            '/* End with rsPAY
                            '/* if the payment to the Agent is made from the commission generated
                            
                    End With
                                        '/* End with Commission

                End If
                                    '/* End if The Payment Type to the Agent is based on the Premium Count
                .MoveNext
                        
        Loop
                        
                                                '/* end if calculation based on Premium
End With

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub saveCOMMISSION()
'On Error GoTo err

    Dim rsPAY As ADODB.Recordset, strPAY As String
    Set rsPAY = New Recordset
    
                    
    strPAY = "SELECT * FROM ALISMAgentsPay ;"
    rsPAY.Open strPAY, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsPAY
        .AddNew
    
        If bCorrectCommission = True Then
                !Agentno = Screen.ActiveForm.cboAgentNo
        ElseIf Trim(rsPREMPAY!Paidto) = "AGENT" Then
                !Agentno = rsPOLICY!Agentno
        ElseIf Trim(rsPREMPAY!Paidto) = "SUPERVISOR" Then
                !Agentno = rsPOLICY!Supervisor
        Else
                !Agentno = rsPOLICY!Agentno
        End If
        
        If bcompleteUNDERWRITING = True Then
                !ReceiptNo = "ISSUED"
                !ReceiptDate = Screen.ActiveForm.txtReceiptDate.Text
                !CurrentPeriod = CurrentPeriod
        Else
                !ReceiptNo = Screen.ActiveForm.txtReceiptNo
                !ReceiptDate = Screen.ActiveForm.txtReceiptDate
                !CurrentPeriod = Screen.ActiveForm.txtAccountingPeriod
        End If
        
        !ProductCode = rsPOLICY!ProductCode
        
        If rsPREMPAY!calcBasedon = "T" Then
            
            If bcompleteUNDERWRITING = True Then
                    !Factor = Val(Screen.ActiveForm.txtexpectedpremium)
                Else: !Factor = Val(Screen.ActiveForm.txtExpectedAmount)
            End If
            
        ElseIf rsPREMPAY!calcBasedon = "P" Then
            !Factor = rsPOLICY!PlanPremium
        ElseIf rsPREMPAY!calcBasedon = "R" Then
            !Factor = rsPOLICY!RiderPremium
        ElseIf rsPREMPAY!calcBasedon = "N" Then
            !Factor = 1
        End If
        
        If rsPREMPAY!generatedfrom = "C" Then
                !Factor = !Amount
        End If
        
        !RateApplied = Val(rsPREMPAY!Percent)
        !PaymentType = rsPREMPAY!PaymentType
        
        If rsPREMPAY!calcBasedon = "T" Then
            If bcompleteUNDERWRITING = True Then
                    !Amount = (rsPREMPAY!Percent / 100) * Val(Screen.ActiveForm.txtUnitCount) * (!Factor / rsPAYMENTMODE!coverperiod)
            Else
                    !Amount = (rsPREMPAY!Percent / 100) * Val(Screen.ActiveForm.txtunitsPaid) * (!Factor / rsPAYMENTMODE!coverperiod)
            End If
        End If
        
        If rsPREMPAY!taxable = "Y" Then
            !TaxRate = Val(rsPREMPAY!TaxRate)
            !taxAmount = (!TaxRate / 100) * !Amount
        Else
            !TaxRate = 0
            !taxAmount = 0
        End If
        
        !JobBriefCount = Val(Screen.ActiveForm.txtUnitCount)
        !Status = "PAID"
        !JobBriefNo = Screen.ActiveForm.cboDocumentNo
        
        .Update
        .Requery

    End With


Exit Sub

err:
    ErrorMessage
End Sub

Public Sub saveBATCHCOMMISSION()
'On Error GoTo err

    Dim rsPAY As ADODB.Recordset, strPAY As String
    Set rsPAY = New Recordset
    
                    
    strPAY = "SELECT * FROM ALISMAgentsPay ;"
    rsPAY.Open strPAY, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsPAY
        .AddNew
    
        If Trim(rsPREMPAY!Paidto) = "AGENT" Then
                !Agentno = rsBATCHPOL!Agentno
        ElseIf Trim(rsPREMPAY!Paidto) = "SUPERVISOR" Then
                !Agentno = rsBATCHPOL!Supervisor
        Else
                !Agentno = rsBATCHPOL!Agentno
        End If
        
        If bcompleteUNDERWRITING = True Then
                !ReceiptNo = "ISSUED"
                !ReceiptDate = !IssueDate
                !CurrentPeriod = CurrentPeriod
        Else
                !ReceiptNo = rsBATCHPOL!ReceiptNo
                !ReceiptDate = rsBATCHPOL!TransactionDate
                !CurrentPeriod = rsBATCHPOL!AccountingPeriod
        End If
        
        !ProductCode = rsBATCHPOL!ProductCode
        
        If rsPREMPAY!calcBasedon = "T" Then
            
        !Factor = Val(rsBATCHPOL!ExpectedPremium)
            
        ElseIf rsPREMPAY!calcBasedon = "P" Then
            !Factor = rsBATCHPOL!PlanPremium
        ElseIf rsPREMPAY!calcBasedon = "R" Then
            !Factor = rsBATCHPOL!RiderPremium
        ElseIf rsPREMPAY!calcBasedon = "N" Then
            !Factor = 1
        End If
        
        If rsPREMPAY!generatedfrom = "C" Then
                !Factor = !Amount
        End If
        
        !RateApplied = Val(rsPREMPAY!Percent)
        !PaymentType = rsPREMPAY!PaymentType
        
        If rsPREMPAY!calcBasedon = "T" Then
            If bcompleteUNDERWRITING = True Then
                    !Amount = (rsPREMPAY!Percent / 100) * Val(rsBATCHPOL!UnitCount) * (!Factor / rsPAYMENTMODE!coverperiod)
            Else
                    !Amount = (rsPREMPAY!Percent / 100) * Val(rsBATCHPOL!UnitsPaid) * (!Factor / rsPAYMENTMODE!coverperiod)
            End If
        End If
        
        If rsPREMPAY!taxable = "Y" Then
            !TaxRate = Val(rsPREMPAY!TaxRate)
            !taxAmount = (!TaxRate / 100) * !Amount
        Else
            !TaxRate = 0
            !taxAmount = 0
        End If
        
        !JobBriefCount = Val(rsBATCHPOL!UnitCount)
        !Status = "PAID"
        !JobBriefNo = rsBATCHPOL!DocumentNo
        
        .Update

    End With


Exit Sub

err:
    ErrorMessage
End Sub

