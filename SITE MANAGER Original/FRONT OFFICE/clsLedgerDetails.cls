VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLedgerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim rsreceipt As ADODB.Recordset, bVAL As Boolean
Dim rsPOLICY As ADODB.Recordset
Dim rsDEP As ADODB.Recordset, strDEP As String

Public Sub getLedgerDetails()
On Error GoTo err
    
        With Screen.ActiveForm
       
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Amount ", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Date   ", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Count ", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Received   ", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Type ", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Suspense", .ListView1.Width / 7

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT  ODASMJobBriefLedger.txtReceiptNo,ODASMJobBriefLedger.receiptAmount, ODASMJobBriefLedger.Receiptdate,ODASMJobBriefLedger.unitcount,ODASMJobBriefLedger.Receivedtodate,ODASMJobBriefLedger.receiptType,ODASMJobBriefLedger.statuscode,ODASMJobBriefLedger.suspenseAccount FROM ODASMJobBriefLedger WHERE DocumentNo =  '" & frmALISMLedgerDetails.cboJobBriefNo.Text & "';"

                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
                        
                        If Not IsNull(rsLIST!ReceiptAmount) Then
                            MyList.SubItems(1) = CStr(rsLIST!ReceiptAmount)
                        End If

                        If Not IsNull(rsLIST!ReceiptDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!ReceiptDate)
                        End If
                        
                        If Not IsNull(rsLIST!UnitCount) Then
                            MyList.SubItems(3) = CStr(rsLIST!UnitCount)
                        End If

                        If Not IsNull(rsLIST!ReceivedToDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!ReceivedToDate)
                        End If

                        If Not IsNull(rsLIST!ReceiptType) Then
                            MyList.SubItems(5) = CStr(rsLIST!ReceiptType)
                        End If

                        If Not IsNull(rsLIST!StatusCode) Then
                            MyList.SubItems(6) = CStr(rsLIST!StatusCode)
                        End If
                        
                        If Not IsNull(rsLIST!SuspenseAccount) Then
                            MyList.SubItems(7) = CStr(rsLIST!SuspenseAccount)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub disableDATAENTRY()
On Error GoTo err

    bVAL = True
    
    Call EnableReceipt
    
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub EnableReceipt()

On Error GoTo err:

       With frmALISMLedgerDetails
            .txtReceivedTodate.Locked = True
            .txtUnitCountBeforePayment.Locked = True
            .txtDueDate.Locked = True
            .txtStatusCode.Locked = True
            .txtNames.Locked = True
            .txtDateOfLastPayment.Locked = True
        End With
    
Exit Sub

err:
    ErrorMessage
End Sub





Public Sub loadPolicy()
On Error GoTo err
            
        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASMJobBrief, ODASPAccount WHERE ODASMJobBrief.JobBriefNo  = '" & frmALISMLedgerDetails.cboJobBriefNo & " ' and ODASMJobBrief.AccountNo = ODASPAccount.AccountNo;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                                     
                If .EOF And Not .BOF Then Exit Sub
                
                 frmALISMLedgerDetails.txtReceivedTodate.Text = !ReceivedToDate & ""
                 frmALISMLedgerDetails.txtpsuspense.Text = !SuspenseAccount & ""
                 frmALISMLedgerDetails.txtDateOfLastPayment.Text = !DateofLastPayment & ""
                 frmALISMLedgerDetails.txtReceivedTodate.Text = !ReceivedToDate & ""
                 frmALISMLedgerDetails.txtUnitCountBeforePayment.Text = !UnitCountBeforePayment & ""
                 frmALISMLedgerDetails.txtDueDate.Text = !DueDate & ""
                 frmALISMLedgerDetails.txtStatusCode.Text = !StatusCode & ""
                 frmALISMLedgerDetails.txtexpectedpremium.Text = !ExpectedPremium & ""
                 frmALISMLedgerDetails.txtDateOfCommencement.Text = !DateOfCommencement & ""
                 frmALISMLedgerDetails.txtPlanPremium.Text = !PlanPremium & ""
                 frmALISMLedgerDetails.txtMaturityDate.Text = !MaturityDate & ""
                 frmALISMLedgerDetails.txtlastpaydate.Text = !DateofLastPayment & ""
                 frmALISMLedgerDetails.txtPaymentMode.Text = !PaymentMode & ""
                 frmALISMLedgerDetails.txtTermOfPolicy.Text = !TermOfPolicy & ""
                 frmALISMLedgerDetails.cboPaymentMethod.Text = !PaymentMethod & ""
                 frmALISMLedgerDetails.txtPremiumDue.Text = !NoofPremiumsdue & ""
                 frmALISMLedgerDetails.txtDateOfLastPayment.Text = !DateofLastPayment & ""
                 frmALISMLedgerDetails.txtUnitCountBeforePayment.Text = !UnitCountBeforePayment & ""
                 frmALISMLedgerDetails.txtDueDate.Text = !DueDate & ""
                 frmALISMLedgerDetails.txtPremiumcount.Text = !UnitCount & ""
                frmALISMLedgerDetails.txtNames = Trim(!OtherNames) + " " + Trim(!CompanyName)
                frmALISMLedgerDetails.txtsex = !GenderCode & ""
                frmALISMLedgerDetails.txtdob = !BirthDate & ""
        End With

rsCONTROL.Close

strSQL = ""
Exit Sub
        
err:         ErrorMessage
        
End Sub




Public Sub LoadNAMES()

        On Error GoTo err
            
            Dim rsREF As ADODB.Recordset, strREF As String
            Set rsREF = New Recordset
            
            strREF = "SELECT * FROM ODASPAccount,ODASMJobBrief Where ODASPAccount.AccountNo=ODASMJobBrief.AccountNo and  ODASMJobBrief.JobBriefNo= '" & frmALISMLedgerDetails.cboJobBriefNo & "';"
            rsREF.Open strREF, cnCOMMON, adOpenKeyset, adLockOptimistic
        
            With rsREF
                If .BOF Or .EOF Then
                            MsgBox "The Policyholder Does not Exist", vbOKOnly
                            Exit Sub
                Else
                    frmALISMLedgerDetails.txtNames = Trim(!OtherNames) + " " + Trim(!CompanyName)
                    frmALISMLedgerDetails.txtsex = !GenderCode & ""
                    frmALISMLedgerDetails.txtdob = !BirthDate & ""
                End If
                
            End With
            rsREF.Close
            
            Dim rsAgent As ADODB.Recordset, strAgent As String
            Set rsAgent = New Recordset
            
            strAgent = "SELECT * FROM ALISPAgent,ODASMJobBrief Where ALISPAgent.AgentNo=ODASMJobBrief.AgentNo and  ODASMJobBrief.JobBriefNo= '" & frmALISMLedgerDetails.cboJobBriefNo & "';"
            rsAgent.Open strAgent, cnCOMMON, adOpenKeyset, adLockOptimistic
        
            With rsAgent
                    If .BOF Or .EOF Then
                            'MsgBox "The agent Does not Exist", vbOKOnly
                            Exit Sub
                    Else
                  frmALISMLedgerDetails.txtagent = Trim(!OtherNames) + " " + Trim(!CompanyName)
                      End If
            End With
            rsAgent.Close
        
        
        Exit Sub
        
err:
                ErrorMessage
End Sub








