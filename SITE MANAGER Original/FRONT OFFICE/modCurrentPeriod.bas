Attribute VB_Name = "modCurrentPeriod"
Option Explicit
Public rsRCPT As ADODB.Recordset, strRCPT, CurrentQuarter, CurrentMonth, UnderwritingYear, CompanyCode, InsurerCode As String
Dim rsPTLF As ADODB.Recordset, strPTLF, StrMonth, strYEAR As String

Public Function CurrentPeriod() As String
    Set rsPTLF = New Recordset
    rsPTLF.Open "SELECT * FROM ODASPDefault; ", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsPTLF
        If .EOF And .BOF Then Exit Function
        CompanyCode = !CompanyCode
        
        StrMonth = Trim(Str(Month(Date)))
        If Len(StrMonth) = 1 Then
            StrMonth = "0" + StrMonth
        End If
        
        CurrentMonth = StrMonth

        strYEAR = Trim(Str(Year(Date)))
        UnderwritingYear = strYEAR
        CurrentPeriod = Trim(strYEAR) + "/" + Trim(StrMonth)
        
        '/ Check whether the Current Period Exist within the Database
        Dim rsPERIOD As New ADODB.Recordset
        Set rsPERIOD = New Recordset
        rsPERIOD.Open "SELECT * FROM ALISPPeriod Where AccountingPeriod = '" & CurrentPeriod & "'; ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPERIOD
                If .BOF Or .EOF Then
                        .AddNew
                        !AccountingPeriod = CurrentPeriod
                        !dateprepared = Date
                        !Preparedby = CurrentUserName
                        !Description = Trim(MonthName(StrMonth)) + " " + strYEAR
                        !AccountMonth = StrMonth
                        !AccountYear = strYEAR
                        !StartDate = Date
                        .Update
                End If

        End With
        
        rsPERIOD.Close
    End With
End Function
