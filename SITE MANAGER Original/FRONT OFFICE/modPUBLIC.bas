Attribute VB_Name = "modPUBLIC"
Option Explicit
'Public strSQL As String
Public MyPassword As String, NewChange As Boolean
Public NewRecord As Boolean, EditRECORD As Boolean
Public rsNewRecord As ADODB.Recordset, rsEditRecord As ADODB.Recordset, rsFindRecord As ADODB.Recordset, rsLineUpdate As ADODB.Recordset, rsCOMBO As ADODB.Recordset
Public PremiumRefundNo As Integer
Public MembersAllNames As String, strMessage As String, AdminPass As String, CurrentRecord As String, CurrentPic As String, INPQRY As Variant, CurrentPrint As String

Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameW Lib "kernel32" (lpBuffer As Any, nSize As Long) As Long

Private StrUserName$
Private strCompName$
Private Res&
Private Les&

'===Discharge Report details

'==Preceeds
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
Public GlobalReferenceNo As Double, GlobalAccidentNo As String
Public loanstype As Boolean, agentsbenefits As Boolean, claimconfig As Boolean, claimcauses As Boolean
Public claimsreq As Boolean, loanapprovers As Boolean, loantype As Boolean, receipts As Boolean, loanoptype As Boolean

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
