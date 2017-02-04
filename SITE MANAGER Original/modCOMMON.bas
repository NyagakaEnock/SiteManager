Attribute VB_Name = "modCOMMON"
Option Explicit

Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameW Lib "kernel32" (lpBuffer As Any, nSize As Long) As Long

Public ThisUser As Variant, LTime As Variant, dtLastDateOfYear As Variant, PaymentsInAYear As Variant, DontChange As Boolean, NewChange As Boolean, ThisSystem As Variant, CLoginID As Variant, rsTestAdmin As ADODB.Recordset, CountLogin As Integer, rsLOGIN As ADODB.Recordset, MyCompName As String, MyLoginID As Variant, QuotationNumber As Variant, JobCardNo As Variant, JobBriefNo As Variant

Public rsLineUpdate, rsBillBoardSchedule As ADODB.Recordset, rsLease1 As ADODB.Recordset, rsLease As ADODB.Recordset, rsFindRecord2 As ADODB.Recordset, rsFindRecord1 As ADODB.Recordset, rsFindRecord4 As ADODB.Recordset, rsFIND As ADODB.Recordset, rsFindRecord As ADODB.Recordset, rsNewRecord As ADODB.Recordset, rsEditRecord As ADODB.Recordset, rsDeleteRecord As ADODB.Recordset, rsBackUpdate As ADODB.Recordset, rsDropDown As ADODB.Recordset, rsCOMBO As ADODB.Recordset

Public strUserName$, strCompName$, Res&, Les&, MyCurrentPeriod As String, MyCurrentMonth As Long, CurrentUserName As String

Public SystemPassword As String, MyPassword As String, Pass2 As String, NewRecord As Boolean, editRECORD As Boolean, PrintDraft As Boolean, MyCurrentDate As String, INPQRY As String, AllowEdit As Boolean, AllowTransferDetails As Boolean, INPQR As String, APCode As Variant

Public MySQL, ActiveTrue As Boolean, NewSQL As String, EditSQL As String, FindSQL As String, DeleteSQL As String, AttachSQL As String

Public cnSECURE As ADODB.Connection, cnCOMMON As ADODB.Connection, cnINVENT As ADODB.Connection, MyCommonData As clsCommonData, i, j, k, l

Public aA, bB, cC, dD, eE, fF, gG, hH, qI, CurrentVisitNo As Variant, MySelectedYear As Variant, UpdateSucceeded As Boolean

Public strSEARCHSQL As String

Public Const vbKeyDecpt = 46
Public PartialPaid As Variant

Public Function GetFullDecryption() As Variant
With Screen.ActiveForm
Dim i, j, k, xx, cC
    j = Len(.txtPassword.Text)
    If j = 0 Then Exit Function
    .txtPassword.SetFocus
    For i = 0 To j - 1
        .txtPassword.SelStart = i: .txtPassword.SelLength = 1: cC = .txtPassword.SelText
        If cC = " " Then
            k = " "
        Else
            k = Trim(GetSetting(appname:="SmallSyzSecure", Section:="SysSecureDecryptor", Key:=cC))
        End If
        xx = xx & k
    Next i
    GetFullDecryption = xx
End With
End Function

Public Function Encryption() As Variant
With Screen.ActiveForm
Dim i, j, k, xx, cC
    j = Len(.txtNewPass.Text)
    If j = 0 Then Exit Function
    .txtNewPass.SetFocus
    For i = 0 To j - 1
        .txtNewPass.SelStart = i: .txtNewPass.SelLength = 1: cC = .txtNewPass.SelText
        If cC = " " Then
            k = " "
        Else
            k = Trim(GetSetting(appname:="SmallSyzSecure", Section:="SysSecureEncryptor", Key:=cC))
        End If
        xx = xx & k
    Next i
    If xx = Empty Then
        Call SaveEncryptionCode
        Call SaveDecryptionCode
        Call Encryption
    Else
        Encryption = xx
    End If
End With
End Function
Public Function GetFullEncryption() As Variant
With Screen.ActiveForm
Dim i, j, k, xx, cC
    j = Len(.txtPassword.Text)
    If j = 0 Then Exit Function
    .txtPassword.SetFocus
    For i = 0 To j - 1
        .txtPassword.SelStart = i: .txtPassword.SelLength = 1: cC = .txtPassword.SelText
        If cC = " " Then
            k = " "
        Else
            k = Trim(GetSetting(appname:="SmallSyzSecure", Section:="SysSecureEncryptor", Key:=cC))
        End If
        xx = xx & k
    Next i
    If xx = Empty Then
        Call SaveEncryptionCode
        Call SaveDecryptionCode
        Call GetFullEncryption
    Else
        GetFullEncryption = xx
    End If
End With
End Function
Public Function FullEncryption() As Variant
With Screen.ActiveForm
Dim i, j, k, xx, cC
    j = Len(.txtCurrentPass.Text)
    If j = 0 Then Exit Function
    .txtCurrentPass.SetFocus
    For i = 0 To j - 1
        .txtCurrentPass.SelStart = i: .txtCurrentPass.SelLength = 1: cC = .txtCurrentPass.SelText
        If cC = " " Then
            k = " "
        Else
            k = Trim(GetSetting(appname:="SmallSyzSecure", Section:="SysSecureEncryptor", Key:=cC))
        End If
        xx = xx & k
    Next i
    If xx = Empty Then
        Call SaveEncryptionCode
        Call SaveDecryptionCode
        Call FullEncryption
    Else
        FullEncryption = xx
    End If
End With
End Function

Public Sub SaveEncryptionCode()
With Screen.ActiveForm
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "a", "!"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "b", "@"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "c", "#"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "d", "$"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "e", "%"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "f", "^"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "g", "&"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "h", "*"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "i", "("
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "j", ")"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "k", "-"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "l", "_"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "m", "="
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "n", "+"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "o", "\"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "p", "|"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "q", "/"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "r", ">"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "s", "<"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "t", "?"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "u", "["
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "v", "]"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "w", "~"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "x", "{"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "y", "}"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "z", ","
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "0", "Z"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "1", "Y"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "2", "X"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "3", "W"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "4", "V"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "5", "U"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "6", "T"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "7", "S"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "8", "R"
    SaveSetting "SmallSyzSecure", "SysSecureEncryptor", "9", "Q"
End With
Exit Sub
End Sub

Public Sub SaveDecryptionCode()
With Screen.ActiveForm
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "!", "a"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "@", "b"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "#", "c"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "$", "d"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "%", "e"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "^", "f"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "&", "g"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "*", "h"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "(", "i"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", ")", "j"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "-", "k"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "_", "l"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "=", "m"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "+", "n"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "\", "o"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "|", "p"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "/", "q"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", ">", "r"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "<", "s"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "?", "t"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "[", "u"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "]", "v"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "~", "w"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "{", "x"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "}", "y"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", ",", "z"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "Z", "0"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "Y", "1"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "X", "2"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "W", "3"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "V", "4"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "U", "5"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "T", "6"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "S", "7"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "R", "8"
    SaveSetting "SmallSyzSecure", "SysSecureDecryptor", "Q", "9"
End With
Exit Sub
End Sub

Public Function GetFullEncryption2() As Variant
With Screen.ActiveForm
Dim i, j, k, xx, cC
    j = Len(.txtConfirmPass.Text)
    If j = 0 Then Exit Function
    For i = 0 To j - 1
        .txtConfirmPass.SelStart = i: .txtConfirmPass.SelLength = 1: cC = .txtConfirmPass.SelText
        If cC = " " Then
            k = " "
        Else
            k = Trim(GetSetting(appname:="SmallSyzSecure", Section:="SysSecureEncryptor", Key:=cC))
        End If
        xx = xx & k
    Next i
    If xx = Empty Then
        Call SaveEncryptionCode
        Call SaveDecryptionCode
        Call GetFullEncryption2
    Else
        GetFullEncryption2 = xx
    End If
End With
End Function

Public Sub AttachDropDowns()
On Error GoTo err
With Screen.ActiveForm
    
    Set rsDropDown = New ADODB.Recordset
    rsDropDown.Open AttachSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If TypeOf Screen.ActiveForm.ActiveControl Is ComboBox Then
        If rsDropDown.EOF And rsDropDown.BOF Then Exit Sub
            rsDropDown.MoveFirst
            Do While Not rsDropDown.EOF
                Do Until rsDropDown!selectfield <> APCode And Not rsDropDown.EOF
                    rsDropDown.MoveNext: If rsDropDown.EOF Then rsDropDown.MoveLast: GoTo OUTS
                Loop
            
                If Not IsNull(rsDropDown!selectfield) Then
                    Screen.ActiveForm.ActiveControl.AddItem rsDropDown!selectfield & ""
                End If
                
                APCode = Trim(rsDropDown!selectfield & "")
            rsDropDown.MoveNext
            Loop
    End If
OUTS:
    APCode = Empty
    Set rsDropDown = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub CheckActiveProcess()
On Error GoTo err
    If NewRecord = True Or editRECORD = True Then
        MsgBox "Requested Operation Cancelled due to Incomplete Transaction! You can Click Refresh to Cancel Ongoing Transactions!", vbCritical + vbOKOnly, "Transaction Monitor"
        ActiveTrue = True
    Else
        ActiveTrue = False
    End If
    Exit Sub
err:
    ErrorMessage
End Sub

Public Function RecordSelected() As Boolean
On Error GoTo err
With Screen.ActiveForm
    If .listView1.ListItems.Count = 0 Or .listView1.View <> lvwReport Or .listView1.Visible = False Then
        
        RecordSelected = False
        
    Else
        
        j = .listView1.ListItems.Count: k = 0
        For i = 1 To j
            If .listView1.ListItems(i).Checked = True Then
                k = k + 1
            End If
        Next i
        
        If k = 0 Then
            RecordSelected = False
        ElseIf k >= 1 Then
            RecordSelected = True
        End If
        
        If Not RecordSelected Then
            MsgBox "No Record Selected to Perform Required Operation...!!", vbCritical + vbOKOnly, "No Selection"
        End If
        
    End If
End With
Exit Function
err:
    ErrorMessage
End Function

Public Function RecordRemoveSelected() As Boolean
On Error GoTo err
With Screen.ActiveForm
    If .listView1.ListItems.Count = 0 Or .listView1.View <> lvwReport Or .listView1.Visible = False Then
        
        RecordRemoveSelected = False
        
    Else
        
        j = .listView1.ListItems.Count: k = 0
        For i = 1 To j
            If .listView1.ListItems(i).Checked = True Then
                k = k + 1
            End If
        Next i
        
        If k = 0 Then
            RecordRemoveSelected = False
        ElseIf k >= 1 Then
            RecordRemoveSelected = True
        End If
        
    End If
End With
Exit Function
err:
    ErrorMessage
End Function

Public Function LoginAccessAllowed() As Boolean
On Error GoTo err
With Screen.ActiveForm
    ThisUser = .txtUserName.Text
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM AdminIndividualRights WHERE UserName='" & ThisUser & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GoTo TestForAdministration
    Else
        GoTo ExamineRights
    End If
    
TestForAdministration:

    Set rsTestAdmin = New ADODB.Recordset
    
    rsTestAdmin.Open "SELECT UserGroup FROM AdminUserRegister WHERE UserName='" & ThisUser & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsTestAdmin.EOF And rsTestAdmin.BOF Then
        LoginAccessAllowed = False: CountLogin = CountLogin + 1
        MsgBox "There are No Data Access Rights Assigned to Your Profile! Contact System Administrator!!", vbOKOnly + vbCritical, "Access Denied": GoTo OUTS
    ElseIf rsTestAdmin!usergroup = "ADMIN" Then
        LoginAccessAllowed = True: GoTo OUTS
    Else
        LoginAccessAllowed = False: CountLogin = CountLogin + 1
        MsgBox "There are No Data Access Rights Assigned to Your Profile! Contact System Administrator!!", vbOKOnly + vbCritical, "Access Denied": GoTo OUTS
    End If
    
ExamineRights:
    'determine the rights due to the user to access current system
    ThisSystem = Trim(App.EXEName)
    
    Select Case ThisSystem
    Case "SYSMGR"
        If rsFindRecord!SYSMGR = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case "JBRIEF"
        If rsFindRecord!JBRIEF = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case "SITEMGR"
        If rsFindRecord!SITEMGR = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case "PURMGR"
        If rsFindRecord!PURMGR = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case "TRIPMST"
        If rsFindRecord!TRIPMST = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case "FOFFICE"
        If rsFindRecord!FOFFICE = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case Else
        LoginAccessAllowed = False: CountLogin = CountLogin + 1
    End Select
    
OUTS:

    If Not LoginAccessAllowed Then
        MsgBox "SORRY!! You do not have the Right to Log-in-To or use this Module!!! Consult Your System Administrator!!!", vbCritical + vbOKOnly, "Module Access Denied"
    End If
    
    Set rsFindRecord = Nothing
    
End With
Exit Function
err:
ErrorMessage
End Function

Public Sub SaveLoginRecord()
On Error GoTo err
    Set rsLOGIN = New ADODB.Recordset
    rsLOGIN.Open "SELECT * FROM AdminUserLog ORDER BY LoginID;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsLOGIN
        .AddNew
        
        !UserName = CurrentUserName
        !LoginDate = Date
        !LoginTime = Format(Now, "hh:mm:ss AMPM")
        !CompName = MyCompName
        !systemused = App.EXEName
                         
        .Update
        
    End With
    
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub UpdateLogoutRecord()
On Error GoTo err
    
    Set rsLOGIN = New ADODB.Recordset
    rsLOGIN.Open "SELECT * FROM UserLog WHERE loginid LIKE '" & MyLoginID & "' ORDER BY LoginID;", cnSECURE, adOpenKeyset, adLockOptimistic

    With rsLOGIN
        If .EOF Or .BOF Then
        Else
            !LogOutDate = Date
            !LogOutTime = Format(Now, "hh:mm:ss AMPM")
    
            .Update
        End If
    End With
    
    Exit Sub
err:
If err.Number = 13 Or err.Number = 3704 Then Resume Next
    ErrorMessage
End Sub

Public Sub OpenODBCConnection()
On Error GoTo err
    Set cnCOMMON = New ADODB.Connection
    
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    GetConfigDetails
    If Database = "" Then Database = "ODAS"
    If Server = "" Then Server = ".\SQLEXPRESS"
    If DSN = "" Then DSN = "ODAS"
    If Uid = "" Then Uid = "sa"
    If Pwd = "" Then Pwd = "Magnate@2010"
    'CreateSQLODBC DSN, Server, Database, "ODAS Plus - Outdoor Advertising Software", Uid, Pwd
    
    Set cnCOMMON = New ADODB.Connection
    cnCOMMON.ConnectionString = "Provider=MSDASQL;DSN=" & DSN & ";UID=" & Uid & ";PWD=" & Pwd & ";"
    cnCOMMON.CommandTimeout = 0
    
    cnCOMMON.Open
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    'cnCOMMON.Open: On Error GoTo err

  
 Exit Sub
err:
    If err.Number = -2147467259 Then
        If MsgBox("The System Cannot Immediately Establish a Connection To the Server! Check to Ensure that Your Server is Running then..." & vbCrLf & "Choose Retry to attemp ODBC Connection, or..." & vbCrLf & "Choose Cancel to Try Connecting With SQL Server OLEDB!!", vbRetryCancel + vbExclamation + vbDefaultButton1, "ODBC Connection Failure... " & App.Title) = vbRetry Then
           cnCOMMON.Cancel: Set cnCOMMON = Nothing
            Call OpenODBCConnection
        Else
            End
        End If
    Else
        ErrorMessage
    End If
End Sub
Public Sub OpenSECUREConnection()
On Error GoTo err
    Set cnSECURE = New ADODB.Connection
    GetConfigDetails
    '5CreateSQLODBC "RIGHTS", Server, "fsecure", "ODAS Plus - Outdoor Advertising Software", Uid, Pwd
    cnSECURE.ConnectionString = "Provider=MSDASQL;DSN=Rights;UID=" & Uid & ";PWD=" & Pwd & ";"
    cnSECURE.CommandTimeout = 0
    
    
    cnSECURE.Open: On Error GoTo err

    Exit Sub
err:
    If err.Number = -2147467259 Then
        If MsgBox("The System Cannot Immediately Establish a Connection To the Server! Check to Ensure that Your Server is Running then..." & vbCrLf & "Choose Retry to attemp ODBC Connection, or..." & vbCrLf & "Choose Cancel to Try Connecting With SQL Server OLEDB!!", vbRetryCancel + vbExclamation + vbDefaultButton1, "ODBC Connection Failure... " & App.Title) = vbRetry Then
           cnSECURE.Cancel: Set cnCOMMON = Nothing
            Call OpenSECUREConnection
        Else
            Call OpenSECUREConnection
        End If
    Else
        ErrorMessage
    End If
End Sub

Public Function getPeriod(TestDate) As String
On Error GoTo err
        Dim strmonth, stryear As String
        strmonth = Trim(Str(Month(TestDate)))
        
        If Len(strmonth) = 1 Then
            strmonth = "0" + strmonth
        End If
        
        stryear = Trim(Str(Year(TestDate)))
        getPeriod = Trim(stryear) + "/" + Trim(strmonth)

Exit Function

err:
    ErrorMessage
End Function

Public Function CurrentPeriod() As String
On Error GoTo err
        Dim strmonth, stryear As String
        strmonth = Trim(Str(Month(Date)))
        
        If Len(strmonth) = 1 Then
            strmonth = "0" + strmonth
        End If
        
        stryear = Trim(Str(Year(Date)))
        CurrentPeriod = Trim(stryear) + "/" + Trim(strmonth)

Exit Function

err:
    ErrorMessage
End Function

Public Sub ErrorMessage()
    If err.Number = 7 Then
        MsgBox "Your System is Short of Memory!" & vbCrLf & "You have too many programs running!" & vbCrLf & "Close some of the programs or Upgrade your computer!", vbInformation + vbOKOnly, "Memory Manager"
    Else
        If err.Number = 20 Then Exit Sub
        MsgBox err.Number & vbCrLf & err.Description, vbInformation, "System Error"
    End If
End Sub

Public Function GetMyLoginID() As String
On Error GoTo err
    Dim rsFindRecord As ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    Dim DTG As String
    DTG = Format(Date, "MMMM dd,yyyy")

    rsFindRecord.Open "SELECT * FROM AdminUserLog WHERE UserName='" & CurrentUserName & "' AND LoginDate='" & DTG & "' ORDER BY LoginID;", cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsFindRecord
    If .EOF And .BOF Then GetMyLoginID = "": GoTo OUTS
        .MoveLast
        GetMyLoginID = !LoginID
    End With
    
OUTS:
    Set rsFindRecord = Nothing
Exit Function
err:
    ErrorMessage
End Function

Public Sub UpdateErrorMessage()
    
    If err.Number = -2147217873 Or err.Number = -2147217900 Then
        MsgBox "The Record Cannot be saved due to Database Primary Key Violation! A similar record already exists in the database!", vbCritical + vbOKOnly, "Canceling Update"
    ElseIf err.Number = -2147467259 Then
        MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical + vbOKOnly, "Canceling Update"
    ElseIf err.Number = -2147352571 Then
        MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical + vbOKOnly, "Canceling Update"
    ElseIf err.Number = -2147217913 Then
        MsgBox "Update Cancelled! Record cannot be saved because of a Date or Other Data-Type Conversi'On Error. Check all Fields to Ensure They Contain Correct Data Types!!!", vbCritical + vbOKOnly, "Data Conversi'On Error"
    Else
        ErrorMessage
    End If
End Sub

Public Function ActiveProcess() As Boolean
On Error GoTo err
    If NewRecord = True Or editRECORD = True Then
        MsgBox "Requested Operation Cancelled due to Incomplete Transaction! You can Click Refresh to Cancel Ongoing Transactions!", vbCritical + vbOKOnly, "Transaction Monitor"
        ActiveProcess = True
    Else
        ActiveProcess = False
    End If
    Exit Function
err:
    ErrorMessage
End Function

Public Function TransactionPeriod() As String
On Error GoTo err

    Dim DateToday As Date
    Dim AccPeriod As String
    
    Dim MonthNow As Variant
    Dim YearNow As Variant
    
        DateToday = Date
        MonthNow = Month(DateToday)
        YearNow = Year(DateToday)
        
        If MonthNow = 10 Or MonthNow = 11 Or MonthNow = 12 Then
            AccPeriod = YearNow & "/" & MonthNow
        Else
            AccPeriod = YearNow & "/" & 0 & MonthNow
        End If
        
        TransactionPeriod = AccPeriod
Exit Function
err:
    ErrorMessage
End Function

Public Function CurrentYear() As Double
On Error GoTo err
    CurrentYear = Year(Date)

Exit Function
err:
    ErrorMessage
End Function

Public Function LeaveYear() As Double
On Error GoTo err
    Dim rsLeave As ADODB.Recordset
    Set rsLeave = New ADODB.Recordset
    
    rsLeave.Open "select * from ParamLVCoyLeaveDefaults;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsLeave
        If .EOF And .BOF Then Exit Function
        LeaveYear = !LeaveYear
    End With
    
Exit Function
err:
    ErrorMessage
End Function

Public Function SysTransactionPeriod() As String
On Error GoTo err
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT SetCurrentPeriod.CurrentPeriod FROM SetCurrentPeriod;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsFindRecord
        If .EOF And .BOF Then
            MsgBox "System Default Period Missing or Empty! Current Month Will be automatically used!!", vbExclamation + vbOKOnly, "Default Period"
            SysTransactionPeriod = TransactionPeriod
        ElseIf IsNull(!CurrentPeriod) = True Or !CurrentPeriod = "" Then
            MsgBox "System Default Period Missing or Empty! Current Month Will be automatically used!!", vbExclamation + vbOKOnly, "Default Period"
            SysTransactionPeriod = TransactionPeriod
        Else
            SysTransactionPeriod = !CurrentPeriod
        End If
    End With
    
    Set rsFindRecord = Nothing
Exit Function
err:
    ErrorMessage
End Function

Public Function LastAccPeriod() As String
On Error GoTo err
    Dim CMonth As Long, LMonth As Long, NMonth As Variant, LPeriod As String, CYear As Variant, NYear As Variant
    
    CYear = CLng(Left(MyCurrentPeriod, 4))
    
    CMonth = CLng(MyCurrentMonth)
    LMonth = CMonth - 1
    
    If LMonth = 0 Then
        NYear = CYear - 1
        NMonth = 12
    Else
        NYear = CYear
        NMonth = LMonth
    End If
    
    If NMonth = 10 Or NMonth = 11 Or NMonth = 12 Then
        LPeriod = NMonth
    Else
        LPeriod = 0 & NMonth
    End If
    
    LastAccPeriod = NYear & "/" & LPeriod
    
Exit Function
err:
    ErrorMessage
End Function

Public Function SysCurrentMonth() As Long
On Error GoTo err
Dim MCPeriod As String, MCM As Variant, MCB As Variant

    MCPeriod = MyCurrentPeriod
    
    MCM = InStr(1, MCPeriod, "/", 1)
    
    MCB = Right(MCPeriod, 2)
    
    SysCurrentMonth = CLng(MCB)
    
Exit Function
err:
    ErrorMessage
End Function

Public Function TransactionMonth() As Long
On Error GoTo err
    Dim DateToday As Date
    Dim AccPeriod As Long
    
    Dim MonthNow As Variant
    Dim YearNow As Variant
    
        DateToday = Date
        MonthNow = Month(DateToday)
        YearNow = Year(DateToday)
        
        AccPeriod = MonthNow
   
        TransactionMonth = AccPeriod
Exit Function
err:
    ErrorMessage
End Function

Public Function RefreshMessage() As String
    RefreshMessage = "This action terminates all ongoing processes! Are you sure you want to continue? All un-saved date will be lost!"
End Function

Public Function PreviousTransPeriod() As String
On Error GoTo err
    Dim CMonth As Long, CYear As Long, SMonth As Long, SYear As Long, NMonth As Long, NYear As Long, PRMonth As String
    
    'determine the current month and year
    CMonth = Month(Date): CYear = Year(Date)
    
    'determine the previous month and year
    SMonth = CMonth - 1
    If SMonth = 0 Then
        NYear = CYear - 1
        NMonth = 12
    Else
        NYear = CYear
        NMonth = SMonth
    End If
    
    'determine the previous month in acc. period format
    If NMonth = 10 Or NMonth = 11 Or NMonth = 12 Then
        PRMonth = NMonth
    Else
        PRMonth = 0 & NMonth
    End If
    
    PreviousTransPeriod = NYear & "/" & PRMonth
    
Exit Function
err:
    ErrorMessage
End Function

Public Function GetCompName() As String
    Dim NullChar&
    
    strCompName = String(1024, 0)
    Les = 1024
    
    Res = GetComputerName(strCompName, Les)
    
    If Res <> 0 Then
        NullChar = InStr(strCompName, vbNullChar)
        strCompName = Trim(Mid(strCompName, 1, NullChar - 1))
        GetCompName = Trim(strCompName)
    End If
End Function

Public Function UserNameGet() As String
    Dim NullChar&
    
    strUserName = String(1024, 0)
    Les = 1024
    
    Res = GetUserName(strUserName, Les)
    
    If Res <> 0 Then
            NullChar = InStr(strUserName, vbNullChar)
            strUserName = Mid(strUserName, 1, NullChar - 1)
            UserNameGet = strUserName
    End If
End Function

Public Function GetModification() As String
    GetModification = "EDIT"
End Function

Public Function GetCurrentStaffDeptCode() As Variant
On Error GoTo err
    Dim rsStaffID As ADODB.Recordset
    Set rsStaffID = cnCOMMON.Execute("SELECT DeptCode FROM ParamEmpMAster WHERE StaffIDNo='" & Trim(GetCurrentStaffID) & "';")
    
    If rsStaffID.EOF And rsStaffID.BOF Then
        GetCurrentStaffDeptCode = Empty
    ElseIf IsNull(rsStaffID!DeptCode) = True Or rsStaffID!DeptCode = "" Then
        GetCurrentStaffDeptCode = Empty
    Else
        GetCurrentStaffDeptCode = rsStaffID!DeptCode & ""
    End If
    
    Set rsStaffID = Nothing
    
Exit Function
err:
    ErrorMessage
End Function

Public Function GetCurrentStaffID() As Variant
On Error GoTo err
    Dim rsStaffID As ADODB.Recordset
    Set rsStaffID = cnCOMMON.Execute("SELECT StaffIDNo FROM AdminUserRegister WHERE UserName='" & Trim(CurrentUserName) & "';")
    
    If rsStaffID.EOF And rsStaffID.BOF Then
        GetCurrentStaffID = Empty
    ElseIf IsNull(rsStaffID!staffidno) = True Or rsStaffID!staffidno = "" Then
        GetCurrentStaffID = Empty
    Else
        GetCurrentStaffID = rsStaffID!staffidno & ""
    End If
    
    Set rsStaffID = Nothing
    
Exit Function
err:
    ErrorMessage
End Function

Public Function GetCurrentStaffFullNames() As Variant
On Error GoTo err
    Dim rsStaffID As ADODB.Recordset
    Set rsStaffID = cnCOMMON.Execute("SELECT AdminUserRegister.StaffIDNo,ParamEmpMAster.Allnames FROM AdminUserRegister,ParamEmpMAster WHERE AdminUSerRegister.StaffIDNo=ParamEmpMaster.StaffIDNo AND AdminUserRegister.UserName='" & Trim(CurrentUserName) & "';")
    
    If rsStaffID.EOF And rsStaffID.BOF Then
        GetCurrentStaffFullNames = Empty
    ElseIf IsNull(rsStaffID!AllNames) = True Or rsStaffID!AllNames = "" Then
        GetCurrentStaffFullNames = Empty
    Else
        GetCurrentStaffFullNames = rsStaffID!AllNames & ""
    End If
    
    Set rsStaffID = Nothing
    
Exit Function
err:
    ErrorMessage
End Function

Public Function GetMyCompanyCode() As String
On Error GoTo err

    Set rsFindRecord = cnCOMMON.Execute("SELECT CompanyCode FROM ParamCompanyMaster WHERE CompanyCode IS NOT NULL;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetMyCompanyCode = Empty
    ElseIf rsFindRecord!CompanyCode = "" Then
        GetMyCompanyCode = Empty
    Else
        GetMyCompanyCode = rsFindRecord!CompanyCode
    End If
        
    Set rsFindRecord = Nothing
    
Exit Function
err:
    ErrorMessage
End Function

Public Function MyCompanyName() As String
On Error GoTo err

    Set rsFindRecord = cnCOMMON.Execute("SELECT CompanyName FROM ParamCompanyMaster WHERE CompanyName IS NOT NULL;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        MyCompanyName = Empty
    ElseIf rsFindRecord!CompanyName = "" Then
        MyCompanyName = Empty
    Else
        MyCompanyName = rsFindRecord!CompanyName
    End If
        
    Set rsFindRecord = Nothing
    
Exit Function
err:
    ErrorMessage
End Function


Public Sub GetConfigDetails(Optional strAltFile As String = "")

Dim strConfigFileName As String
Dim strBackSlash As String
Dim intConfigFileNbr As Integer
 
Dim strConfigLn As String
Dim strConfigSetting

strBackSlash = IIf(Right$(App.Path, 1) = "\", "", "\")
If strAltFile = "" Then
    strConfigFileName = App.Path & strBackSlash & "config.DAT"
Else
    strConfigFileName = App.Path & strBackSlash & "configALT.DAT"
End If
Dim fso As New FileSystemObject
If fso.FileExists(strConfigFileName) Then
        intConfigFileNbr = FreeFile
        
        Open strConfigFileName For Input As #intConfigFileNbr
        
        Do Until EOF(intConfigFileNbr)
            Input #intConfigFileNbr, strConfigLn
            'Debug.Print strConfigLn;
            strConfigSetting = Split(strConfigLn, ":=")
            Select Case strConfigSetting(0)
                    Case "Database"
                            Database = strConfigSetting(1)
                    Case "Dsn"
                            DSN = strConfigSetting(1)
                    Case "server"
                            Server = strConfigSetting(1)
                    Case "uid"
                            Uid = strConfigSetting(1)
                    Case "pwd"
                            Pwd = strConfigSetting(1)
            End Select
        Loop
Else
        Database = ""
        Server = ""
        Uid = ""
        Pwd = ""
End If
End Sub

