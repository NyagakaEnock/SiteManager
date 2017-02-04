VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMoneyConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub CallConverter()

On Error GoTo err
    Dim FNum As Double
    'remove redundant characters
    FNum = figures
    figures = Trim(Replace(FNum, ",", ""))
    figures = FormatNumber(CDbl(figures), 2, vbTrue, vbUseDefault, vbFalse)
    Call MoneyConverter
    Exit Sub

err:
    If err.Number = 13 Then
        MsgBox "Enter a numeric value to convert!", vbExclamation, "Format Number"
        'Me.txtWtemp.SetFocus
        Exit Sub
    Else
        MsgBox err.Number & vbCrLf & err.Description, vbExclamation, "Format Number"
    End If
End Sub

Public Sub MoneyConverter()

'''On Error Resume Next
Dim tempModulars, txtTotalLength As Double
Dim WCents, WPendingString As String
Dim WPendoriginal, WAllowPassage, Wstring2, Wstring3, Wstring1, Wcount, WString As String
Dim LTemp, LpreviousTemp, LNextTemp, L2NextTemp, LthirdTemp As Double
Dim LStr As Integer
Dim txtWDivide, txtWNextDivide, WPass As Double
       
            figures = Trim(figures)
            Wcount = InStr(figures, ".") - 1
            txtTotalLength = Len(figures)
           WPendingString = "N"
           WAllowPassage = "N"
           WPass = 0
           LTemp = 0
           txtWDivide = 0
           txtWNextDivide = 0
           tempModulars = 0
           LthirdTemp = 0
           LpreviousTemp = 0
           LNextTemp = 0
           L2NextTemp = 0
           WAllowPassage = "N"
           inwords = ""
  
 Do
           Wstring2 = ""
           Wstring3 = ""
           WString = ""
           WAllowPassage = "N"
           
            WPass = CDbl(WPass) + 1
            
            LTemp = Left(Trim(figures), WPass)
            LTemp = LTemp Mod 10
            
            txtWDivide = (Wcount - 1) / 3
            txtWDivide = Left(txtWDivide, 1)
            tempModulars = (Wcount - 1) Mod 3
 
            If WPass > 1 And Wcount > 2 Then
            
                            If WPass >= 3 Then
                                        LthirdTemp = Left(Trim(figures), WPass - 2)
                                        LthirdTemp = LthirdTemp Mod 10
                            End If
                            
                            LpreviousTemp = Left(Trim(figures), WPass - 1)
                            LpreviousTemp = LpreviousTemp Mod 10
                            LNextTemp = Left(Trim(figures), WPass + 1)
                            LNextTemp = LNextTemp Mod 10
                            
                            txtWNextDivide = (Wcount - 2) / 3
                            txtWNextDivide = Left(txtWNextDivide, 1)
                            
                            If tempModulars = 0 And (txtWNextDivide <> txtWDivide) And (LpreviousTemp > 0 Or LNextTemp > 0 Or LthirdTemp > 0) Then
                                    WAllowPassage = "Y"
                            End If
            End If
            
     
     If (tempModulars = 0 And (LTemp > 0 Or WAllowPassage = "Y")) Then
                Select Case txtWDivide
                Case 0
                            WString = ""
                Case 1
                            WString = "Thousand"
                Case 2
                            WString = "Million"
                Case 3
                            WString = "Billion"
                Case 4
                            WString = "Trillion"
                End Select
        End If
        
    
         If (tempModulars = 2 Or tempModulars = 0) And WPendingString = "N" Then
                Select Case LTemp
                        Case 1
                                    Wstring2 = "ONE"
                        Case 2
                                    Wstring2 = "TWO"
                        Case 3
                                    Wstring2 = "Three"
                        Case 4
                                    Wstring2 = "Four"
                        Case 5
                                    Wstring2 = "Five"
                        Case 6
                                    Wstring2 = "SIX"
                        Case 7
                                    Wstring2 = "Seven"
                        Case 8
                                    Wstring2 = "Eight"
                        Case 9
                                    Wstring2 = "Nine"
                        Case 0
                                    Wstring2 = ""
                End Select
        End If



            If WPendingString = "Y" Then
                    Select Case LTemp
                        Case 1
                                    Wstring2 = "Eleven"
                        Case 2
                                    Wstring2 = "Twelve"
                        Case 3
                                    Wstring2 = "Thirteen"
                        Case 4
                                    Wstring2 = "Fourteen"
                        Case 5
                                    Wstring2 = "Fifteen"
                        Case 6
                                    Wstring2 = "Sixteen"
                        Case 7
                                    Wstring2 = "Seventeen"
                        Case 8
                                    Wstring2 = "Eighteen"
                        Case 9
                                    Wstring2 = "Nineteen"
                        Case 0
                                    Wstring2 = "Ten"
                    End Select
                    
                    WPendingString = "N"
                End If
 
            
            If tempModulars = 1 And WPendingString = "N" Then
                Select Case LTemp
                        Case 1
                                    Wstring2 = ""
                                    WPendingString = "Y"
                        Case 2
                                    Wstring2 = "Twenty"
                        Case 3
                                    Wstring2 = "Thirty"
                        Case 4
                                    Wstring2 = "Fourty"
                        Case 5
                                    Wstring2 = "Fifty"
                        Case 6
                                    Wstring2 = "Sixty"
                        Case 7
                                    Wstring2 = "Seventy"
                        Case 8
                                    Wstring2 = "Eighty"
                        Case 9
                                    Wstring2 = "Ninety"
                        Case 0
                                    Wstring2 = ""
                End Select
            End If
                
    '/***************************************************************/
    '/* This pocedure is used to print one hundred when the */
    '/* modulars is 2 elase the string remails blank                */
    '/***************************************************************/
            LNextTemp = Left(Trim(figures), WPass + 1)
            LNextTemp = LNextTemp Mod 10
            
            L2NextTemp = Left(Trim(figures), WPass + 2)
            L2NextTemp = L2NextTemp Mod 10
            
            If WPass > 1 Then
                    LpreviousTemp = Left(Trim(figures), WPass - 1)
                    LpreviousTemp = LpreviousTemp Mod 10
            End If
 
            Select Case tempModulars
            Case 2
            If LTemp > 0 Then
                If LNextTemp = 0 And L2NextTemp = 0 Then
                         Wstring3 = "Hundred"
                 End If
                 
                 If LNextTemp > 0 And L2NextTemp = 0 Then
                         Wstring3 = "Hundred and"
                 End If
                 
                 If LNextTemp = 0 And L2NextTemp > 0 Then
                         Wstring3 = "Hundred and"
                 End If
                 
                 If LNextTemp > 0 And L2NextTemp > 0 Then
                         Wstring3 = "Hundred and"
                 End If
                        
            ElseIf LTemp = 0 And WPass > 1 Then
                If LNextTemp = 0 And L2NextTemp > 0 Then
                        Wstring3 = Wstring3 + "AND"
                End If
                
                If LNextTemp > 0 And L2NextTemp > 0 Then
                        Wstring3 = Wstring3 + "AND"
                End If
                
                If LNextTemp > 0 And L2NextTemp = 0 Then
                        Wstring3 = Wstring3 + "AND"
                End If
                
                If LNextTemp = 0 And L2NextTemp = 0 Then
                        Wstring3 = ""
                End If
                        
            Else: Wstring3 = ""
            End If
                           
            Case 1
                Wstring3 = ""
            Case 0
                Wstring3 = ""
                                        
      End Select
    
    
   '/****************************************************************/
   '/* This procedure is used to concatenate several strings */
   '/* generated above                                                               */
   '/****************************************************************/

    Wstring2 = Trim(Wstring2)
    Wstring3 = Trim(Wstring3)
    WString = Trim(WString)
    
    If tempModulars <> 0 Then
            inwords = inwords + " " + Wstring2 + " " + Wstring3
            Else: inwords = inwords + " " + Wstring2 + " " + Wstring3 + " " + WString
    End If
    
    inwords = Trim(inwords)
    
    '/****************************************************************/
    '/* The next Procedure is used to reduce the word count  */
    '/* by 1 and repeat the procedure above                             */
    '/****************************************************************/

    Wcount = Wcount - 1
    
    Loop Until Wcount = 0
    
    
    '/*************************************************************************************
    '/ This procedure is used to Translate the Decimal part of the number        */
    '/* entered by the user                                                                                         */
   '/*************************************************************************************
            LNextTemp = Right(Trim(figures), 2)
            LNextTemp = Left(LNextTemp / 10, 1)
            
            
            LTemp = Right(Trim(figures), 1)
            LTemp = LTemp Mod 10

            Wstring2 = ""
            Wstring3 = ""
            WCents = ""
            
        Select Case LNextTemp
            Case 1
                Wstring2 = ""
                WPendingString = "Y"
            Case 2
                Wstring2 = "Twenty"
            Case 3
                Wstring2 = "Thirty"
            Case 4
                Wstring2 = "Fourty"
            Case 5
                Wstring2 = "Fifty"
            Case 6
                Wstring2 = "Sixty"
            Case 7
                Wstring2 = "Seventy"
            Case 8
                Wstring2 = "Eighty"
            Case 9
                Wstring2 = "Ninety"
            Case 0
                Wstring2 = ""
        End Select
                
                
        If WPendingString = "Y" Then
        Select Case LTemp
            Case 1
                Wstring2 = "Eleven"
            Case 2
                Wstring2 = "Twelve"
            Case 3
                Wstring2 = "Thirteen"
            Case 4
                Wstring2 = "Fourteen"
            Case 5
                Wstring2 = "Fifteen"
            Case 6
                Wstring2 = "Sixteen"
            Case 7
                Wstring2 = "Seventeen"
            Case 8
                Wstring2 = "Eighteen"
            Case 9
                Wstring2 = "Nineteen"
            Case 0
                Wstring2 = "Ten"
        End Select
        End If
            
        If LNextTemp <> 1 Then
        Select Case LTemp
            Case 1
                Wstring3 = "ONE"
            Case 2
                Wstring3 = "TWO"
            Case 3
                Wstring3 = "Three"
            Case 4
                Wstring3 = "Four"
            Case 5
                Wstring3 = "Five"
            Case 6
                Wstring3 = "SIX"
            Case 7
                Wstring3 = "Seven"
            Case 8
                Wstring3 = "Eight"
            Case 9
                Wstring3 = "Nine"
            Case 0
                Wstring3 = ""
           End Select
        End If
            
      If LTemp = 0 And LNextTemp = 0 Then
            WCents = ""
      End If
                                                
      If LTemp > 0 And LNextTemp = 0 Then
            WCents = "and " + Wstring3 + " Cents"
      End If
                                                
      If LTemp = 0 And LNextTemp > 0 Then
            WCents = "and " + Wstring2 + " Cents"
      End If
                                                
      If LTemp > 0 And LNextTemp > 0 Then
            WCents = "and " + Wstring2 + " " + Wstring3 + " Cents"
      End If
             
    If inwords = "" Then
            inwords = "NILL"
        Else
            inwords = inwords + " " + WCents + " Only"
        End If
        
        inwords = UCase(inwords)
          
End Sub
