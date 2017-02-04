Attribute VB_Name = "modPUBLIC1"
Option Explicit
Public strPAYEE, strPOSTALCODE, strPOSTALADDRESS, strTOWN, strRCPT As String, LTime As Variant
Public MyPassword, strCheckNo As String, NewChange As Boolean
Public NewRecord, bschedulePAYMENT, bscheduledCHEQUES, baddBENEFICIARY As Boolean, rsINVOICE, rsJOBBRIEF As ADODB.Recordset
Public rsNewRecord, rsCONTROL, rsPaymentCode, rsDEFAULT As ADODB.Recordset, rsEditRecord, rsSAVE As ADODB.Recordset, rsFindRecord As ADODB.Recordset, rsLineUpdate As ADODB.Recordset, rsCOMBO As ADODB.Recordset
Public bapprovePROPOSAL, bapprovePOLICY, bauthorizePOLICY, bauthorizePROPOSAL, bPolicyAPPROVAL, bpendingPROPOSAL, bapprovedPROPOSAL As Boolean
Public PremiumRefundNo, df As Integer
Public MembersAllNames, GlobalApplicationNo, GlobalOperationType, GlobalOperationDescription, GlobalClaimNo As String, strMessage As String, AdminPass As String, CurrentRecord As String, CurrentPic As String, INPQRY As Variant, CurrentPrint As String
Public figures, strTEMP As String
Public inwords As String

Public CurrentRecord2, CurrentRecord1 As Variant
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameW Lib "kernel32" (lpBuffer As Any, nSize As Long) As Long

Private StrUserName$
Private strCompName$
Private Res&
Private Les&

Public SelectedListItem As String

Public Edit As Boolean

Public paidupsump As Double, addsump As Double, bonusp As Double, pmannuityp As Double
Public suspensep As Double, surrenderp As Double, interimbonusp As Double
Public anticipated As Double
'==Deductions
Public Premiumdued As Double, outloand As Double, outinterestd As Double, penaltyd As Double
'===Other details
Public sumassured As Double, claimno As String, netpayable As Double, premiumduedated As String
Public proceeds As Double, deductions As Double, premiumrefundp As Double
Public surrendervaluep As Double, Preparedby As String, dateprepared As String
Public agentnames As String, Agentno As String
Public suspense As Double, penalty As Double, unpaidpremium As Double, loaninterest As Double, loan As Double
'===For Parameter Reports
Public tittles As Boolean, paymethods As Boolean, taxes As Boolean, employers As Boolean, cities As Boolean
Public countries As Boolean, accperiods As Boolean, currencies As Boolean, feesservices As Boolean, payinterval As Boolean
Public companydepts As Boolean, companybranch As Boolean, companymaster As Boolean
Public moffice As Boolean, enquiry As Boolean, empmaster As Boolean
Public agents As Boolean, agentspay As Boolean, claimletters As Boolean, claimcorrespondence As Boolean
Public claimantdetails As Boolean, surrendersetup As Boolean, periodssetup As Boolean, companydetails As Boolean
Public ratetable As Boolean, jointage As Boolean, bankssetup As Boolean, lastnumbers As Boolean
Public claimdischarge As Boolean
Public GlobalAccountNo As Double, GlobalAccidentNo As String
Public loanstype As Boolean, agentsbenefits As Boolean, claimconfig As Boolean, claimcauses As Boolean
Public claimsreq As Boolean, loanapprovers As Boolean, loantype As Boolean, receipts As Boolean, loanoptype As Boolean
Public Function RefreshMessage() As String
    RefreshMessage = "This action terminates all ongoing processes! Are you sure you want to continue? All un-saved date will be lost!"
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
Public Sub GetCharacterEncryptionCode()
With Screen.ActiveForm
Dim CODE As Variant
CODE = Trim(.txtPassword.Text)
    .txtPassword.Text = Trim(GetSetting(appname:="SmallSyzSecure", Section:="SysSecureEncryptor", Key:=CODE))
End With
End Sub

Public Sub calcTotalPremium()
On Error GoTo err

        Screen.ActiveForm.txtexpectedpremium = CDbl(Screen.ActiveForm.txtPlanPremium) + CDbl(Screen.ActiveForm.txtRiderPremium)
        
        '/ Obtain The Deposit Protection Fund - PolicyHolder Contribution
        
        Set rsCONTROL = New ADODB.Recordset
    
        strSQL = "SELECT * FROM ALISPDefaults ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.BOF Or rsCONTROL.EOF Then Exit Sub
        Screen.ActiveForm.txtDepositProtectionFund.Text = CDbl(Screen.ActiveForm.txtexpectedpremium) * CDbl(rsCONTROL!DPF_PHContribution)
        
        Wtemp = Screen.ActiveForm.txtDepositProtectionFund.Text
        'RoundTo1Shilling
        Screen.ActiveForm.txtDepositProtectionFund.Text = Wtemp
        Screen.ActiveForm.txtexpectedpremium = CDbl(Screen.ActiveForm.txtexpectedpremium) + CDbl(Screen.ActiveForm.txtDepositProtectionFund.Text)
Exit Sub

err:
    ErrorMessage
End Sub

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
    
    StrUserName = String(1024, 0)
    Les = 1024
    
    Res = GetUserName(StrUserName, Les)
    
    If Res <> 0 Then
            NullChar = InStr(StrUserName, vbNullChar)
            StrUserName = Mid(StrUserName, 1, NullChar - 1)
            UserNameGet = StrUserName
    End If
End Function

Public Function GetModification() As String
    GetModification = "EDIT"
End Function
