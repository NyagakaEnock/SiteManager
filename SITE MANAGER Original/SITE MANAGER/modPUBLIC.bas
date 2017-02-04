Attribute VB_Name = "modPUBLIC"
Option Explicit

Public Const vbKeyDecpt = 46

Public MyPassword As String, NewChange As Boolean
Public NewData As Boolean, EditData As Boolean
Public rsMEDIA As ADODB.Recordset, rsDEFAULT As ADODB.Recordset
Public AppPath As String, RetVal As Variant
Public MembersAllNames As String, strMessage As String, AdminPass As String, CurrentRecord, CurrentStartDate, CurrentEndDate, CurrentRecord1 As Variant, CurrentPic As String, CurrentPrint As String, INPQRY2 As Variant

Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameW Lib "kernel32" (lpBuffer As Any, nSize As Long) As Long

Private StrUserName$
Private strCompName$
Private Res&
Private Les&

Public ExistEntry As Boolean

Public Database As String, Server As String, DSN As String, Pwd As String, Uid As String

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

Public Function GetModification() As String
    GetModification = "EDIT"
End Function

Public Function GetCurrentStaffID() As Variant
On Error GoTo err
    Set rsFindRecord = cnCOMMON.Execute("SELECT StaffIDNo FROM AdminUserRegister WHERE UserName='" & Trim(CurrentUserName) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetCurrentStaffID = Empty
    Else
        GetCurrentStaffID = rsFindRecord!staffidno & ""
    End If
    
    Set rsFindRecord = Nothing
    
Exit Function
err:
    ErrorMessage
End Function

Public Function SystemActivated() As Boolean
On Error GoTo err
    Dim i, j
    Select Case App.EXEName
    Case "SYSMGR"
        j = "Systems0"
    Case "SITEMGR"
        j = "Systems1"
    Case "TRIPMST"
        j = "Systems2"
    Case "PURMGR"
        j = "Systems3"
    Case "FOFFICE"
        j = "Systems4"
    Case "JBRIEF"
        j = "Systems5"
    Case "JCARD"
        j = "Systems6"
    Case "ADMIN"
        j = "Systems7"
    Case Else
        j = Empty: SystemActivated = False: Exit Function
    End Select
    
    i = GetSetting(appname:="XAdmin", Section:="XActive", Key:=j)
    
    If i = "" Then
        SystemActivated = False
    ElseIf i = 1 Then
        SystemActivated = True
    ElseIf i = 0 Then
        SystemActivated = False
    End If
    
    If Not SystemActivated Then
        MsgBox "This Application Cannot Run as Of Now Because it Has Not Been Activated!! Kindly Contact the Copyright Owners for the Activation to be Effected!", vbCritical + vbOKOnly, "Mandatory System Activation"
    End If
    
    Exit Function
err:
    ErrorMessage
End Function


Public Sub CreateSQLODBC(txtDSN As String, txtServer As String, txtDatabase As String, txtDescription As String, Optional txtUid As String = "", Optional txtPwd As String = "", Optional OptSQL As Boolean = True)
' This function setups a DSN common for remote Database servers
' Such as SQL or Oracle, keep in mind, this isnt a complete listing of parameters
Dim RetVal

   Dim DSN           As String
   Dim Server        As String
   Dim Address       As String
   Dim Database      As String
   Dim Description   As String
   Dim Uid           As String
   Dim Pwd           As String
   
   Dim Security      As String
   Dim SqlDriver     As String
   Dim SQLParameter As String
   'Basically the DSN Name you want to have
   DSN = "DSN=" & Trim(txtDSN)
   'The IP Addy of the server you want , if this is a remote connection
   Server = "SERVER=" & Trim(txtServer)
   'Same as above
   Address = "ADDRESS=" & Trim(txtServer)
   'The name of the database as known by the DB Server, such as SQL Server
   Database = "DATABASE=" & Trim(txtDatabase)
   'An Optional Description Feild
   Description = "DESCRIPTION=" & Trim(txtDescription)
   'This is optional, if you require a Security mode check the help files
   Uid = "UID=" & txtUid
   Pwd = "PWD=" & txtPwd
   Security = "NETWORK=dbmssocn"
   
   'the next couple lines setup the Driver Text , that defines the type of DB Drivers
   ' you are using, if its anything other than the ones I've listed, check your DB
   ' documentation, or check the ODBC settings to see it's names
      
   'Also you will notice as each string peice is put together, they are seperated by
   'VbNullChar, this gives it a Null seperated array in a sense so that the API Command
   'can use the Parameters
   
   If OptSQL = True Then
      SqlDriver = "SQL Server"
      SQLParameter = DSN & vbNullChar & Server & vbNullChar & Address & vbNullChar & Security & vbNullChar & _
         Database & vbNullChar & Description & vbNullChar & Uid & vbNullChar & Pwd & vbNullChar & vbNullChar
   Else
      SqlDriver = "Oracle73"
      SQLParameter = DSN & vbNullChar & Server & vbNullChar & Database & vbNullChar & _
         Description & vbNullChar & Uid & vbNullChar & Pwd & vbNullChar & vbNullChar
   End If
      
   'calls SQLConfigDataSource , giving it the forms handle, the command to Add a System DSN
   'giving it the Driver name, and then the Null Seperated Parameter listing
    Dim tDSNDetails As tDSNAttrib, sError As String
    With tDSNDetails
        .Database = txtDatabase
        .Driver = SqlDriver
        .Server = txtServer
        .TrustedConnection = True    'Use NT authentication
        .Password = txtPwd
        .UserID = txtUid
        .DSN = txtDSN
        .Description = txtDescription
        .Type = ServerBased
        .SystemDSN = False           'Create a System DSN
    End With

    sError = DSNCreate(tDSNDetails)
'   retVal = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, SqlDriver, SQLParameter)
   'Replace 0& with Me.hwnd if you wish for users to further configure settings such as a long
   If Len(sError) > 0 Then
      MsgBox sError, vbExclamation
    End If
End Sub

