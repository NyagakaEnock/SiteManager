Attribute VB_Name = "modConnections"
Option Explicit
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public AttachSQL As Variant, NewSQL, EditSQL, FindSQL, DeleteSQL, UpdateSucceeded As Boolean
Public cnCOMMON As ADODB.Connection, cnSECURE As New ADODB.Connection, cnREGISTRY As ADODB.Connection, rsDropDown As ADODB.Recordset
Public cnPAY As ADODB.Connection, CurrentUserName
Public Wtemp As Double, CLoginID, DateIssued As Variant

Public Function GetCurrentStaffDeptCode() As Variant
On Error GoTo err
    Dim rsStaffID As ADODB.Recordset
    Set rsStaffID = cnREGISTRY.Execute("SELECT DeptCode FROM ParamEmpMAster WHERE StaffIDNo='" & Trim(GetCurrentStaffID) & "';")
    
    If rsStaffID.EOF And rsStaffID.BOF Then
        GetCurrentStaffDeptCode = Empty
    ElseIf IsNull(rsStaffID!deptcode) = True Or rsStaffID!deptcode = "" Then
        GetCurrentStaffDeptCode = Empty
    Else
        GetCurrentStaffDeptCode = rsStaffID!deptcode & ""
    End If
    
    Set rsStaffID = Nothing
    
Exit Function
err:
    ErrorMessage
End Function

Public Function GetCurrentStaffID() As Variant
On Error GoTo err
    Dim rsStaffID As ADODB.Recordset
    Set rsStaffID = cnREGISTRY.Execute("SELECT StaffIDNo FROM AdminUserRegister WHERE UserName='" & Trim(CurrentUserName) & "';")
    
    If rsStaffID.EOF And rsStaffID.BOF Then
        GetCurrentStaffID = Empty
    ElseIf IsNull(rsStaffID!StaffIdNo) = True Or rsStaffID!StaffIdNo = "" Then
        GetCurrentStaffID = Empty
    Else
        GetCurrentStaffID = rsStaffID!StaffIdNo & ""
    End If
    
    Set rsStaffID = Nothing
    
Exit Function
err:
    ErrorMessage
End Function

Public Function GetCurrentStaffFullNames() As Variant
On Error GoTo err
    Dim rsStaffID As ADODB.Recordset
    Set rsStaffID = cnREGISTRY.Execute("SELECT AdminUserRegister.StaffIDNo,ParamEmpMAster.Allnames FROM AdminUserRegister,ParamEmpMAster WHERE AdminUSerRegister.StaffIDNo=ParamEmpMaster.StaffIDNo AND AdminUserRegister.UserName='" & Trim(CurrentUserName) & "';")
    
    If rsStaffID.EOF And rsStaffID.BOF Then
        GetCurrentStaffFullNames = Empty
    ElseIf IsNull(rsStaffID!allnames) = True Or rsStaffID!allnames = "" Then
        GetCurrentStaffFullNames = Empty
    Else
        GetCurrentStaffFullNames = rsStaffID!allnames & ""
    End If
    
    Set rsStaffID = Nothing
    
Exit Function
err:
    ErrorMessage
End Function

Public Sub OpenConnection()
On Error GoTo err

    Set cnCOMMON = New ADODB.Connection
    
    cnCOMMON.ConnectionString = "Provider=MSDASQL;DSN=ODAS;UID=sa;PWD=;"
    cnCOMMON.CommandTimeout = 0
    
     cnCOMMON.Open: On Error GoTo err

    Exit Sub
err:
    If err.Number = -2147467259 Then
        If MsgBox("The System Cannot Immediately Establish a Connection To the A.L.I.S. Server! Check to Ensure that Your Server is Running then Choose Retry to Connect!!", vbRetryCancel + vbExclamation + vbDefaultButton1, "ODBC A.L.I.S. Connection Failure") = vbRetry Then
           cnCOMMON.cancel: Set cnCOMMON = Nothing
            Call OpenConnection
        Else
            MsgBox "This Action terminates the application! Try Loading Again Later!!", vbExclamation + vbOKOnly, "Cancel Connection"
            End
        End If
    Else
        ErrorMessage
    End If
End Sub

Public Sub OpenSECUREConnection()
On Error GoTo err

    Set cnSECURE = New ADODB.Connection
    
    cnSECURE.ConnectionString = "Provider=MSDASQL;DSN=RIGHTS;UID=sa;PWD=;"
    cnSECURE.CommandTimeout = 0
    
     cnSECURE.Open: On Error GoTo err

    Exit Sub
err:
    If err.Number = -2147467259 Then
        If MsgBox("The System Cannot Immediately Establish a Connection To the A.L.I.S. Server! Check to Ensure that Your Server is Running then Choose Retry to Connect!!", vbRetryCancel + vbExclamation + vbDefaultButton1, "ODBC A.L.I.S. Connection Failure") = vbRetry Then
           cnSECURE.cancel: Set cnCOMMON = Nothing
            Call OpenSECUREConnection
        Else
            MsgBox "This Action terminates the application! Try Loading Again Later!!", vbExclamation + vbOKOnly, "Cancel Connection"
            End
        End If
    Else
        ErrorMessage
    End If
End Sub
