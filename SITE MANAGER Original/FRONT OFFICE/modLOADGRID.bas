Attribute VB_Name = "modLOADGRID"
Public Sub loadAGENTTYPEGRID()
On Error GoTo ERR

    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    StrGRID = "SELECT * from ALISPAgentType;"
    rsGRID.Open StrGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISPCreditor.AgentTypeGrid.DataSource = rsGRID

Exit Sub

ERR:
    ErrorMessage
End Sub
Public Sub loadAPPROVERGRID()
On Error GoTo ERR

    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    StrGRID = "SELECT * from ALISPLoanOperationType;"
    rsGRID.Open StrGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISPLoanOperationType.OperationTypeGrid.DataSource = rsGRID

Exit Sub

ERR:
    ErrorMessage
End Sub

Public Sub loadAGENTUNITGRID()
On Error GoTo ERR

    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    StrGRID = "SELECT * from ALISPAgent where unitCode = '" & Screen.ActiveForm.cboUnitCode.Text & "' and STATUS = 'A';"
    rsGRID.Open StrGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISPCreditor.AgentUNITGRID.DataSource = rsGRID

Exit Sub

ERR:
    ErrorMessage
End Sub

Public Sub loadAGENTBRANCHGRID()
On Error GoTo ERR

    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    StrGRID = "SELECT * from ALISPAgent where BranchCode = '" & Screen.ActiveForm.cboBranchCode.Text & "' and STATUS = 'A' ;"
    rsGRID.Open StrGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISPCreditor.AgentBranchGRID.DataSource = rsGRID

Exit Sub

ERR:
    ErrorMessage
End Sub

Public Sub loadAGENTCoyGRID()
On Error GoTo ERR

    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    StrGRID = "SELECT * from ALISPAgent WHERE STATUS = 'A';"
    rsGRID.Open StrGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISPCreditor.AgentCoyGRID.DataSource = rsGRID

Exit Sub

ERR:
    ErrorMessage
End Sub

Public Sub LoadSurnameGRID()
On Error GoTo ERR
        Dim strSurname As String
        Dim rsSurname As ADODB.Recordset

        Set rsSurname = New ADODB.Recordset
        
        strSurname = " Select * from ALISPAgent Where Surname = '" & frmALISPCreditor.txtSurname & "'"
        rsSurname.Open strSurname, cnALIS, adOpenKeyset, adLockOptimistic
        
        If rsSurname.BOF Or rsSurname.EOF Then Exit Sub
            
         Set frmALISPCreditor.SurnameGRID.DataSource = rsSurname
        
Exit Sub


ERR:
    ErrorMessage
End Sub

Public Sub LoadAgentSurnameGRID()
On Error GoTo ERR

        Dim strSurname As String
        Dim rsSurname As ADODB.Recordset

        Set rsSurname = New ADODB.Recordset
        
        strSurname = " Select * from ALISPAgent Where Surname = '" & frmALISPCreditor.txtSurname & "'"
        rsSurname.Open strSurname, cnALIS, adOpenKeyset, adLockOptimistic
        
        If rsSurname.BOF Or rsSurname.EOF Then Exit Sub
            
        Set frmALISPCreditor.SurnameGRID.DataSource = rsSurname
        
Exit Sub


ERR:
    ErrorMessage
End Sub

Public Sub LoadAgentOtherNamesGRID()
On Error GoTo ERR
        Dim strOtherNames As String
        Dim rsOtherNames As ADODB.Recordset

        Set rsOtherNames = New ADODB.Recordset
        
        strOtherNames = " Select * from ALISPAgent Where OtherNames = '" & frmALISPCreditor.txtOtherNames & "'"
        rsOtherNames.Open strOtherNames, cnALIS, adOpenKeyset, adLockOptimistic
        
        If rsOtherNames.BOF Or rsOtherNames.EOF Then Exit Sub
            
         Set frmALISPCreditor.SurnameGRID.DataSource = rsOtherNames
        
Exit Sub

ERR:
    ErrorMessage
End Sub

