Attribute VB_Name = "modLISTVIEWS"
Public Sub showALLCOUNCILS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Council Code", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Council", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Use Calendar Year?", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT CouncilCode, Council, UseCalendarYear, Status FROM ODASPCouncil Where Status = 'A' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                DF = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CouncilCode))
                        
                        If Not IsNull(rsLIST!Council) Then
                            MyList.SubItems(1) = CStr(rsLIST!Council)
                        End If
                        
                        If Not IsNull(rsLIST!UseCalendarYear) Then
                            MyList.SubItems(2) = CStr(rsLIST!UseCalendarYear)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                            MyList.SubItems(3) = CStr(rsLIST!Status)
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

Public Sub ListALLCOUNCILACCOUNTS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView5.ListItems.Clear
                .ListView5.ColumnHeaders.Clear
                
                .ListView5.ColumnHeaders.Add , , "Account No", .ListView5.Width / 3 ', lvwColumnCenter
                .ListView5.ColumnHeaders.Add , , "Company Name", .ListView5.Width / 3
                .ListView5.ColumnHeaders.Add , , "Town", .ListView5.Width / 3
                .ListView5.ColumnHeaders.Add , , "Type", .ListView5.Width / 3

                .ListView5.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPAccount A, ODASPAccountType T Where A.AccountType = T.AccountType and A.status = 'A' and T.Council = 'Y'"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                DF = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView5.ListItems.Add(, , CStr(rsLIST!AccountNo))
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!Towncity) Then
                            MyList.SubItems(2) = CStr(rsLIST!Towncity)
                        End If
                        
                        If Not IsNull(rsLIST!AccountType) Then
                            MyList.SubItems(3) = CStr(rsLIST!AccountType)
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

Public Sub ListALLSITES()
On Error GoTo err
    
        With frmODASMCouncilRates
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 3.5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "B.Board No", .ListView1.Width / 3.5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Site", .ListView1.Width / 2

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlotSite, ODASPPlot Where  ODASPPlotSite.PlotNo = ODASPPlot.PlotNo and ODASPPlot.CouncilCode = '" & .txtTownCode.Text & "' and ODASPPlot.OnRoadReserve='Y';"   ' AND ODASPPlotSite.JobBriefItemNo is not null;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                    
                    
                    If rsLIST.EOF And rsLIST.BOF Then GoTo Continue
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!MastNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!MastNo)
                        End If
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(2) = CStr(rsLIST!PlotName)
                        End If
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                            MyList.SubItems(3) = CStr(rsLIST!SiteDetails)
                        End If
Continue:
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub ShowSITESWITHRATES()
On Error GoTo err
    
        With frmODASMCouncilRates
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Site No", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Media code", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Media Size", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Payment Mode", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Start Date", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "End Date", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Amount", .ListView3.Width / 4
                .ListView3.View = lvwReport
                
                Dim rsLIST, rslist1 As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset: Set rslist1 = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMCouncilRateDue R, ODASPPlotSite S where S.SiteNo = R.SiteNo and (R.AmountDue > 0) ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                DF = rsLIST.RecordCount
                
                Dim MyList As ListItem
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!SiteNo))
                            
                            If Not IsNull(rsLIST!MediaCode) Then
                                MyList.SubItems(1) = CStr(rsLIST!MediaCode)
                            End If

                            If Not IsNull(rsLIST!MediaSize) Then
                                MyList.SubItems(2) = CStr(rsLIST!MediaSize)
                            End If
                            If Not IsNull(rsLIST!PaymentMode) Then
                                MyList.SubItems(3) = CStr(rsLIST!PaymentMode)
                            End If
                            
                            If Not IsNull(rsLIST!StartDate) Then
                                MyList.SubItems(4) = CStr(rsLIST!StartDate)
                            End If
                            
                            If Not IsNull(rsLIST!EndDate) Then
                                MyList.SubItems(5) = CStr(rsLIST!EndDate)
                            End If

                            If Not IsNull(rsLIST!AmountDue) Then
                                MyList.SubItems(6) = CStr(rsLIST!AmountDue)
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


'Public Sub listAPPROVALTASKS()
'On Error GoTo err
'
'        With Screen.ActiveForm
'
'                .ListView1.ListItems.Clear
'                .ListView1.ColumnHeaders.Clear
'                .ListView1.ColumnHeaders.Add , , "User", .ListView1.Width / 5
'                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 5
'                .ListView1.ColumnHeaders.Add , , "Operation Date ", .ListView1.Width / 5
'                .ListView1.ColumnHeaders.Add , , "Comment", .ListView1.Width / 5
'                .ListView1.ColumnHeaders.Add , , "Accept", .ListView1.Width / 5
'
'
'                .ListView1.View = lvwReport
'
'                Dim rsLIST As ADODB.Recordset
'                Set rsLIST = New ADODB.Recordset
'
'                rsLIST.Open "SELECT UserCode, Status, OperationDate, Comment, Accept FROM ODASMOperation WHERE ApplicationNo =  '" & Screen.ActiveForm.txtApplicationNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
'
'                DF = rsLIST.RecordCount
'                Dim MyList As ListItem
'
'                While Not rsLIST.EOF
'
'                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!UserCode))
'
'                        If Not IsNull(rsLIST!Status) Then
'                            MyList.SubItems(1) = CStr(rsLIST!Status)
'                        End If
'
'                        If Not IsNull(rsLIST!operationDate) Then
'                                MyList.SubItems(2) = CStr(rsLIST!operationDate)
'                        End If
'
'                        If Not IsNull(rsLIST!Comment) Then
'                                MyList.SubItems(3) = CStr(rsLIST!Comment)
'                        End If
'
'                        If Not IsNull(rsLIST!Accept) Then
'                                MyList.SubItems(4) = CStr(rsLIST!Accept)
'                        End If
'
'                        rsLIST.MoveNext
'                Wend
'                Set MyList = Nothing
'        End With
'
'Exit Sub
'
'err:
'        If err.Number = 3265 Then Resume Next
'         ErrorMessage
'End Sub
Public Sub showALLPROPERTIES()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Property", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 1


                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open "SELECT * FROM ODASPProperties WHERE Status =  'A';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                        Set rsCONTROL = New ADODB.Recordset
                        strCONTROL = "SELECT * FROM ODASMSiteProperties WHERE sITENo = '" & frmODASPAssignProperties.txtSiteNo.Text & "' and PropertyCode = '" & rsLIST!PropertyCode & "'"
                        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
        

                        If rsCONTROL.EOF Or rsCONTROL.BOF Then
                                Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PropertyCode))
                                
                                If Not IsNull(rsLIST!PropertyDescription) Then
                                    MyList.SubItems(1) = CStr(rsLIST!PropertyDescription)
                                End If
        
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

Public Sub showALLPROPERTIES1()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListALLProperties.ListItems.Clear
                .ListALLProperties.ColumnHeaders.Clear
                .ListALLProperties.ColumnHeaders.Add , , "Property", .ListALLProperties.Width / 5
                .ListALLProperties.ColumnHeaders.Add , , "Description", .ListALLProperties.Width / 1


                .ListALLProperties.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open "SELECT * FROM ODASPProperties WHERE Status =  'A';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                        Set rsCONTROL = New ADODB.Recordset
                        strCONTROL = "SELECT * FROM ODASMSiteProperties WHERE SiteNo = '" & .txtSiteNo.Text & "' and PropertyCode = '" & rsLIST!PropertyCode & "'"
                        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
        

                        If rsCONTROL.EOF Or rsCONTROL.BOF Then
                                Set MyList = .ListALLProperties.ListItems.Add(, , CStr(rsLIST!PropertyCode))
                                
                                If Not IsNull(rsLIST!PropertyDescription) Then
                                    MyList.SubItems(1) = CStr(rsLIST!PropertyDescription)
                                End If
        
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

Public Sub showALLProperties2()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                .ListView2.ColumnHeaders.Add , , "Property", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Description", .ListView2.Width / 1


                .ListView2.View = lvwReport
                
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open "SELECT * FROM ODASPProperties WHERE Status =  'A' and PropertyCode != '" & frmODASSitesProperties.txtFirstPropertyCde.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                        Set rsCONTROL = New ADODB.Recordset
                        strCONTROL = "SELECT * FROM ODASMSiteProperties WHERE SiteNo = '" & frmODASPAssignProperties.txtSiteNo.Text & "' and PropertyCode = '" & rsLIST!PropertyCode & "'"
                        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
        

                        If rsCONTROL.EOF Or rsCONTROL.BOF Then
                                Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!PropertyCode))
                                
                                If Not IsNull(rsLIST!PropertyDescription) Then
                                    MyList.SubItems(1) = CStr(rsLIST!PropertyDescription)
                                End If
        
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

Public Sub showALLMASTS()
On Error GoTo err
    
        With frmODASMSiteRegistration
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                .ListView2.ColumnHeaders.Add , , "Mast No", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Plot No", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Mast Details", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Rent", .ListView2.Width / 4


                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASPPlotMast  WHERE ODASPPlotMast.PlotNo = '" & .txtPlotNo.Text & "' "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                DF = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!MastNo))
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                        End If

                        If Not IsNull(rsLIST!MastDetails) Then
                                MyList.SubItems(2) = CStr(rsLIST!MastDetails)
                        End If
                        
                        If Not IsNull(rsLIST!AnnualRent) Then
                                MyList.SubItems(3) = CStr(rsLIST!AnnualRent)
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
Public Sub showUNALLOCATEDSites()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView4.ListItems.Clear
                .ListView4.ColumnHeaders.Clear
                .ListView4.ColumnHeaders.Add , , "Site No", .ListView4.Width / 5
                .ListView4.ColumnHeaders.Add , , "Plot No", .ListView4.Width / 5
                .ListView4.ColumnHeaders.Add , , "Mast No ", .ListView4.Width / 5
                .ListView4.ColumnHeaders.Add , , "Details", .ListView4.Width / 5
                .ListView4.ColumnHeaders.Add , , "Size", .ListView4.Width / 5

                .ListView4.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlotSite, ODASPPlot WHERE ODASPPlotSite.PlotNo = ODASPPlot.PlotNo AND ODASPPlotSite.Status = 'UNALLOCATED'"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                        Set MyList = .ListView4.ListItems.Add(, , CStr(rsLIST!PlotNo))
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                        End If

                        If Not IsNull(rsLIST!MastNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!MastNo)
                        End If
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                                MyList.SubItems(3) = CStr(rsLIST!SiteDetails)
                        End If
                        
                        If Not IsNull(rsLIST!MediaSize) Then
                                MyList.SubItems(4) = CStr(rsLIST!MediaSize)
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

Public Sub showALLPLOTS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Name", .ListView1.Width / 3.5
                .ListView1.ColumnHeaders.Add , , "LR No ", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Location", .ListView1.Width / 3.5
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 7


                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                If bsearchRECORD = True Then
                    strSQL = "SELECT * FROM ODASPPlot P, ODASPAccount A where A.AccountNo = P.AccountNo and P.TownCode = '" & .txtTownCode.Text & "' AND P.PhysicalLocation like '%" & Trim(.txtPhysicalAddress.Text) & "%' "
                Else: strSQL = "SELECT * FROM ODASPPlot P, ODASPAccount A where A.AccountNo = P.AccountNo and P.TownCode = '" & .txtTownCode.Text & "' "
                End If
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                        
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotName)
                        End If

                        If Not IsNull(rsLIST!LRNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!LRNo)
                        End If
                        
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(3) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!TownCode) Then
                                MyList.SubItems(4) = CStr(rsLIST!TownCode)
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

Public Sub ShowAllWorksDueForMaintenanceSpecificDate()
On Error GoTo err
    
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "MaintananceNo", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4.5
                .ListView1.ColumnHeaders.Add , , "Customer", .ListView1.Width / 5.5
                .ListView1.ColumnHeaders.Add , , "Media", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Maitanance Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                Dim Datex As Date
                Datex = InputBox("Enter  the first Date...", "Lower Date Interval")
                If IsDate(Datex) Then
                    MsgBox "Either the cancel was canceled or no Date was entered", vbCritical + vbOKOnly, "Missing Date"
                    Exit Sub
                Else
                rsLIST.Open "SELECT * FROM ODASPTown T, ODASPPlot P,ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPPlotSite PS,ODASPAccount A,ODASPmedia ME,ODASMMaintenance MN WHERE JBI.JobBriefItemNo = PS.JobBriefItemNo and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and MN.MaintenanceDate = '" & Format(Datex, "MMMM dd,yyyy") & "' and MN.SiteNo = PS.SiteNo and PS.PlotNo = P.PlotNo and P.TownCode = T.TownCode and MN.Maintained = 'N';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                DF = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MaintenanceNo))
                        
                        If Not IsNull(rsLIST!SiteNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteNo)
                        End If

                        If Not IsNull(rsLIST!SiteDetails) Then
                                MyList.SubItems(2) = CStr(rsLIST!SiteDetails)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(3) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!MediaDescription) Then
                                MyList.SubItems(4) = CStr(rsLIST!MediaDescription)
                        End If
                        If Not IsNull(rsLIST!MaintananceDueDate) Then
                                MyList.SubItems(5) = CStr(rsLIST!MaintananceDueDate)
                        End If
                        If Not IsNull(rsLIST!Town) Then
                                MyList.SubItems(6) = CStr(rsLIST!Town)
                        End If
                    rsLIST.MoveNext
                Wend
                Set MyList = Nothing
            End If
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub
Public Sub ShowAllWorksDueForMaintenanceSpcificPeriod()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4.5
                .ListView1.ColumnHeaders.Add , , "Customer", .ListView1.Width / 5.5
                .ListView1.ColumnHeaders.Add , , "Media", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Maitanance Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                Dim Date1, Date2 As Variant
                Date1 = InputBox("Enter  the first Date...", "Lower Date Interval")
                Date2 = InputBox("Enter the second Date...", "Upper Date Interval")
                
                If Len(Date1) = 0 Or Len(Date2) = 0 Then
                MsgBox "Either of the Input request boxes was canceled or neither of the dates were entered", vbCritical + vbOKOnly, "Canceled action/Missing values"
                Exit Sub
                Else
                
                rsLIST.Open "SELECT * FROM ODASPTown T, ODASPPlot P,ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPPlotSite PS,ODASPAccount A,ODASPmedia ME,ODASMMaintenance M WHERE JBI.JobBriefItemNo = PS.JobBriefItemNo and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and M.SiteNo = PS.SiteNo and (M.MaintenanceDate between '" & Format(Date1, "MMMM dd,yyyy") & "' and '" & Format(Date2, "MMMM dd,yyyy") & "') and PS.PlotNo = P.PlotNo and P.TownCode = T.TownCode and M.Maintained = 'N';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                        End If

                        If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!MediaDescription) Then
                                MyList.SubItems(3) = CStr(rsLIST!MediaDescription)
                        End If
                        
                        If Not IsNull(rsLIST!MaintananceDueDate) Then
                                MyList.SubItems(4) = CStr(rsLIST!MaintananceDueDate)
                        End If
                         If Not IsNull(rsLIST!Town) Then
                                MyList.SubItems(5) = CStr(rsLIST!Town)
                        End If
                    rsLIST.MoveNext
                Wend
                Set MyList = Nothing
            End If
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub ShowAllWorksDueForMaintenanceONEMonth()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4.5
                .ListView1.ColumnHeaders.Add , , "Customer", .ListView1.Width / 5.5
                .ListView1.ColumnHeaders.Add , , "Media", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Maitanance Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open "SELECT * FROM ODASMMaintenance M,ODASPTown T, ODASPPlot P,ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPPlotSite PS,ODASPAccount A,ODASPmedia ME WHERE JBI.JobBriefItemNo = PS.JobBriefItemNo and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and JBI.MaintananceDueDate> '" & Format(Date, "MMMM dd,yyyy") & "' and JBI.MaintananceDueDate = M.MaintenanceDate and M.Maintained = 'N'and PS.PlotNo = P.PlotNo and P.TownCode = T.TownCode;", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        If DateDiff("M", Date, rsLIST!MaintananceDueDate) > 1 Then Exit Sub
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                        End If

                        If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!MediaDescription) Then
                                MyList.SubItems(3) = CStr(rsLIST!MediaDescription)
                        End If
                        
                        If Not IsNull(rsLIST!MaintananceDueDate) Then
                                MyList.SubItems(4) = CStr(rsLIST!MaintananceDueDate)
                        End If
                         If Not IsNull(rsLIST!Town) Then
                                MyList.SubItems(5) = CStr(rsLIST!Town)
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

Public Sub showAllJobsCompleted()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "JobBrief ItemNo", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "SiteNo", .ListView1.Width / 6
'                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4.5
                .ListView1.ColumnHeaders.Add , , "Customer", .ListView1.Width / 5.5
                .ListView1.ColumnHeaders.Add , , "Media", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Maitanance Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "BB", .ListView1.Width / 5
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open "SELECT * FROM ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPAccount A,ODASPmedia ME WHERE JB.Closed = 'Y' and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and JBI.ExpiryDate > '" & Format(Date, "MMMM dd,yyyy") & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefItemNo))
                     If Not IsNull(rsLIST!SiteNo) Then
                        MyList.SubItems(1) = CStr(rsLIST!SiteNo)
                    End If
                    If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                    End If
                    
                    If Not IsNull(rsLIST!MediaDescription) Then
                            MyList.SubItems(3) = CStr(rsLIST!MediaDescription)
                    End If
                    
                    If Not IsNull(rsLIST!MaintananceDueDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!MaintananceDueDate)
                    End If
                    If Not IsNull(rsLIST!Town) Then
                            MyList.SubItems(5) = CStr(rsLIST!Town)
                    End If
                    If Not IsNull(rsLIST!BillBoard) Then
                            MyList.SubItems(6) = CStr(rsLIST!BillBoard)
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
Public Sub ShowAllWorksDueForMaintenance()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4.5
                .ListView1.ColumnHeaders.Add , , "Customer", .ListView1.Width / 5.5
                .ListView1.ColumnHeaders.Add , , "Media", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Maitanance Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open "SELECT * FROM ODASPTown T, ODASPPlot P,ODASMJobBriefItems JBI,ODASMJobBrief JB,ODASPPlotSite PS,ODASPAccount A,ODASPmedia ME WHERE JBI.JobBriefItemNo = PS.JobBriefItemNo and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo and JBI.MediaCode = ME.MediaCode and JBI.MaintananceDueDate > '" & Format(Date, "MMMM dd,yyyy") & "' and PS.PlotNo = P.PlotNo and P.TownCode = T.TownCode;", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                        End If

                        If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!MediaDescription) Then
                                MyList.SubItems(3) = CStr(rsLIST!MediaDescription)
                        End If
                        
                        If Not IsNull(rsLIST!MaintananceDueDate) Then
                                MyList.SubItems(4) = CStr(rsLIST!MaintananceDueDate)
                        End If
                         If Not IsNull(rsLIST!Town) Then
                                MyList.SubItems(5) = CStr(rsLIST!Town)
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
Public Sub showMaintenancePROPERTIES()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Property", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Require Payment ", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Calc Due Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "DOC Required", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open "SELECT * FROM ODASPProperties WHERE Status =  'A' and Maintenance = 'Y' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PropertyCode))
                        
                        If Not IsNull(rsLIST!PropertyDescription) Then
                            MyList.SubItems(1) = CStr(rsLIST!PropertyDescription)
                        End If

                        If Not IsNull(rsLIST!RequirePayment) Then
                                MyList.SubItems(2) = CStr(rsLIST!RequirePayment)
                        End If
                        
                        If Not IsNull(rsLIST!CalculateDueDate) Then
                                MyList.SubItems(3) = CStr(rsLIST!CalculateDueDate)
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
Public Sub ShowAllUnApprovedContracts()
On Error GoTo err
With Screen.ActiveForm
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Item Code", .ListView1.Width / 6#
.ListView1.ColumnHeaders.Add , , "Media Type", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Media Name", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Length", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Width", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Quantity Quoted", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "SidingDescription", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Illiminated", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Bordered", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Price", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Total Price", .ListView1.Width / 5.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertQuotationItems A,AdvertBBDetails B,AdvertSiding C WHERE A.ItemCode = B.SiteNo AND A.SidingType = C.SidingType AND A.QuotationNo = '" & QuotationNumber & "'", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ItemCode))


    If Not IsNull(rsLIST!CategoryName) Then
        MyList.SubItems(1) = CStr(rsLIST!CategoryName)
    End If
     
    If Not IsNull(rsLIST!ItemName) Then
        MyList.SubItems(2) = CStr(rsLIST!ItemName)
    End If
    
    If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(3) = CStr(rsLIST!Length)
    End If
    
    If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(4) = CStr(rsLIST!Width)
    End If
    
    If Not IsNull(rsLIST!quantity) Then
        MyList.SubItems(5) = CStr(rsLIST!quantity)
    End If
    
    If Not IsNull(rsLIST!SidingDescription) Then
        MyList.SubItems(6) = CStr(rsLIST!SidingDescription)
    End If
    
    If Not IsNull(rsLIST!illuminated) And (rsLIST!illuminated) = 0 Then
        MyList.SubItems(7) = CStr("NO")
      ElseIf Not IsNull(rsLIST!illuminated) And (rsLIST!illuminated) = 1 Then
        MyList.SubItems(7) = CStr("YES")
    End If
    
    If Not IsNull(rsLIST!BorderType) And (rsLIST!BorderType) = 0 Then
        MyList.SubItems(8) = CStr("NO")
      ElseIf Not IsNull(rsLIST!BorderType) And (rsLIST!BorderType) = 1 Then
        MyList.SubItems(8) = CStr("YES")
    End If
    
    If Not IsNull(rsLIST!Price) Then
        MyList.SubItems(9) = "Ksh" + " " + FormatNumber((rsLIST!Price), 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!TotalPrice) Then
        MyList.SubItems(10) = "Ksh" + " " + FormatNumber((rsLIST!TotalPrice), 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
       
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub showALLAvailableSites()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView4.ListItems.Clear
                .ListView4.ColumnHeaders.Clear
                
                .ListView4.ColumnHeaders.Add , , "Site No", .ListView4.Width / 3 ', lvwColumnCenter
                .ListView4.ColumnHeaders.Add , , "Site Name", .ListView4.Width / 3
                .ListView4.ColumnHeaders.Add , , "Physical Address", .ListView4.Width / 3
                .ListView4.ColumnHeaders.Add , , "Site Details", .ListView4.Width / 3


                .ListView4.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
            
                strSQL = "SELECT ODASPPlotSite.SiteNo, ODASPPlot.PlotName, ODASPPlot.PhysicalLocation, ODASPPlotSite.SiteDetails FROM ODASPPlot, ODASPPlotSite where ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.Status = 'SITE-AVAILABLE' "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set MyList = .ListView4.ListItems.Add(, , CStr(rsLIST!SiteNo))
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                    MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!SiteDetails) Then
                                    MyList.SubItems(3) = CStr(rsLIST!SiteDetails)
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
Public Sub showALLAvailableFaces()
On Error GoTo err
    
        With Screen.ActiveForm
        
               .ListView1.ListItems.Clear
               .ListView1.ColumnHeaders.Clear
                
               .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 9 ', lvwColumnCenter
               .ListView1.ColumnHeaders.Add , , "BillBoard No", .ListView1.Width / 9
               .ListView1.ColumnHeaders.Add , , "Free From", .ListView1.Width / 9
               .ListView1.ColumnHeaders.Add , , "Till", .ListView1.Width / 9
               .ListView1.ColumnHeaders.Add , , "Free Days", .ListView1.Width / 11
               .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 9
               .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4
               .ListView1.ColumnHeaders.Add , , "Plot Location", .ListView1.Width / 3


               .ListView1.View = lvwReport
                
                    Screen.MousePointer = vbHourglass
                    Set rsFindRecord1 = New ADODB.Recordset
                    rsFindRecord1.Open "Select * From ODASPPlot,ODASPTown,ODASPPlotSite,ODASPPlotMast Where ODASPTown.Town like '" & CurrentRecord & "%' and ODASPPlot.PlotNo = ODASPPlotMast.PlotNo and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotMast.ExpiryDate>'" & Format(Date, "MMMM dd,YYYY") & "' and ODASPPlot.TownCode = ODASPTown.TownCode", cnCOMMON, adOpenKeyset, adLockOptimistic
                    If rsFindRecord1.EOF And rsFindRecord1.BOF Then Exit Sub
                    .ProgressBar1.Visible = True: .ProgressBar1.Value = 0: .ProgressBar1.Min = 0: .ProgressBar1.Max = rsFindRecord1.RecordCount
                    rsFindRecord1.MoveFirst
                    Do While rsFindRecord1.EOF <> True
                        Set rsFindRecord = New ADODB.Recordset
                            rsFindRecord.Open "Select min(scheduleDate) as StartDate, max(scheduleDate)as EndDate from ODASMSiteSchedule Where SiteNo  = '" & rsFindRecord1!SiteNo & "' and (Reserved = 'N' or JobBriefItemNo is null) and ScheduleDate >'" & Format(Date, "MMMM dd,YYYY") & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
                            If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo Continue
                     Dim MyList As ListItem
                           
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsFindRecord1!SiteNo))
                            
                            If Not IsNull(rsFindRecord1!MastNo) Then
                                MyList.SubItems(1) = CStr(rsFindRecord1!MastNo)
                            End If
                            
                            If Not IsNull(rsFindRecord!StartDate) Then
                                    MyList.SubItems(2) = CStr(rsFindRecord!StartDate)
                            End If
                            
                            If Not IsNull(rsFindRecord!EndDate) Then
                                    MyList.SubItems(3) = CStr(rsFindRecord!EndDate)
                            End If
                            DF = DateDiff("d", rsFindRecord!StartDate, rsFindRecord!EndDate)
                            If Not IsNull(DF) Then
                                    MyList.SubItems(4) = CDbl(DF)
                            End If
                            If Not IsNull(rsFindRecord1!Town) Then
                                    MyList.SubItems(5) = CStr(rsFindRecord1!Town)
                            End If
                            
                            If Not IsNull(rsFindRecord1!SiteDetails) Then
                                    MyList.SubItems(6) = CStr(rsFindRecord1!SiteDetails)
                            End If
                            If Not IsNull(rsFindRecord1!PhysicalLocation) Then
                                    MyList.SubItems(7) = CStr(rsFindRecord1!PhysicalLocation)
                            End If
Continue:
                         rsFindRecord1.MoveNext:  .ProgressBar1.Value = .ProgressBar1.Value + 1
                     Loop
                Set MyList = Nothing: Screen.MousePointer = vbwindowstate
             .ProgressBar1.Visible = False
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetUserCode()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Staff", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "User Name", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "All Names", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT AdminUserRegister.StaffIdNo,AdminUserRegister.UserName, AdminUserRegister.AllNames FROM AdminUserRegister ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set rsCONTROL = New ADODB.Recordset
                
                        strSQL = "SELECT * FROM ODASPApprovers Where OperationType = '" & Screen.ActiveForm.cboOperationType & "' and StaffId = '" & rsLIST!staffidno & "';"
                        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                        If rsCONTROL.BOF Or rsCONTROL.EOF Then
                                
                                Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!staffidno))
                                    If Not IsNull(rsLIST!UserName) Then
                                        MyList.SubItems(1) = CStr(rsLIST!UserName)
                                    End If
                                    
                                    If Not IsNull(rsLIST!AllNames) Then
                                            MyList.SubItems(2) = CStr(rsLIST!AllNames)
                                    End If
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

Public Sub showDepartmentACCESS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "Staff", .ListView2.Width / 3 ', lvwColumnCenter
                .ListView2.ColumnHeaders.Add , , "User Name", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "All Names", .ListView2.Width / 3

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPDeptAccess, AdminUserRegister where ODASPDeptAccess.Status = 'A' and ODASPDeptAccess.StaffIdNo = AdminUserRegister.StaffIdno and ODASPDeptAccess.DepartmentCode = '" & frmODASPDeptAccess.txtDepartmentCode.Text & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                                
                    Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!staffidno))
                        If Not IsNull(rsLIST!UserName) Then
                            MyList.SubItems(1) = CStr(rsLIST!UserName)
                        End If
                        
                        If Not IsNull(rsLIST!AllNames) Then
                                MyList.SubItems(2) = CStr(rsLIST!AllNames)
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
Public Sub showStaffACCESS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Staff", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "User Name", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "All Names", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT AdminUserRegister.StaffIdNo,AdminUserRegister.UserName, AdminUserRegister.AllNames FROM AdminUserRegister ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                DF = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set rsCONTROL = New ADODB.Recordset
                
                        strCONTROL = "SELECT * FROM ODASPDeptAccess Where StaffIDNo = '" & rsLIST!staffidno & "' and DepartmentCode = '" & frmODASPDeptAccess.txtDepartmentCode.Text & "';"
                        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                        If rsCONTROL.BOF Or rsCONTROL.EOF Then
                                
                                Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!staffidno))
                                    If Not IsNull(rsLIST!UserName) Then
                                        MyList.SubItems(1) = CStr(rsLIST!UserName)
                                    End If
                                    
                                    If Not IsNull(rsLIST!AllNames) Then
                                            MyList.SubItems(2) = CStr(rsLIST!AllNames)
                                    End If
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
Public Sub loadMoreDETAILS()
On Error GoTo err
    With frmODASMAllocation
        Set rsFindRecord2 = New ADODB.Recordset
        rsFindRecord2.Open "SELECT * FROM ODASPAccount A,ODASPPlot P, ODASMLeaseAgreement LA WHERE LA.ContractNo = '" & .txtContractNo.Text & "' and LA.PlotNo = P.PlotNo and LA.AccountNo = A.AccountNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
        .cboPaymentMode.Text = rsFindRecord2!PaymentMode
        .txtAgreementDate.Text = rsFindRecord2!AgreementDate
        .txtSignedBy.Text = rsFindRecord2!SignedBy & ""
        .txtWitnessCoy.Text = rsFindRecord2!WitnessCoy & ""
        .txtWitnessLandLord.Text = rsFindRecord2!WitnessLandLord
        .txtLandLordNo.Text = rsFindRecord2!AccountNo
    End With
Exit Sub
err:
ErrorMessage
End Sub
Public Sub loadMastDETAILS()
On Error GoTo err
    With frmODASMAllocation
        Set rsFindRecord2 = New ADODB.Recordset
        rsFindRecord2.Open "SELECT * FROM ODASPAccount A,ODASPPlot P, ODASMLeaseAgreement LA WHERE LA.ContractNo = '" & .txtContractNo.Text & "' and LA.PlotNo = P.PlotNo and LA.AccountNo = A.AccountNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
        .cboPaymentMode.Text = rsFindRecord2!PaymentMode
        .txtAgreementDate.Text = rsFindRecord2!AgreementDate
        .txtSignedBy.Text = rsFindRecord2!SignedBy & ""
        .txtWitnessCoy.Text = rsFindRecord2!WitnessCoy & ""
        .txtWitnessLandLord.Text = rsFindRecord2!WitnessLandLord
        .txtLandLordNo.Text = rsFindRecord2!AccountNo
    End With
Exit Sub
err:
ErrorMessage
End Sub

Public Sub getLeasableMasts()
On Error GoTo err
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Structure No", .ListView3.Width / 3 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , " Media", .ListView3.Width / 3
                .ListView3.ColumnHeaders.Add , , "Size", .ListView3.Width / 3
 
                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlotMast Where PlotNo = '" & .txtPlotNo.Text & "' AND OwenedByClient = 'N' and (LeasePrepared ='N' or LeasePrepared is null);"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                    Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!MastNo))
                    If Not IsNull(rsLIST!TypeOfMast) Then
                        MyList.SubItems(1) = CStr(rsLIST!TypeOfMast)
                    End If
                    
                    If Not IsNull(rsLIST!MediaSize) Then
                            MyList.SubItems(2) = CStr(rsLIST!MediaSize)
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
Public Sub getMastsToLease()
On Error GoTo err
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Structure No", .ListView3.Width / 3 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , " Media", .ListView3.Width / 3
                .ListView3.ColumnHeaders.Add , , "Size", .ListView3.Width / 3
 
                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlotMast Where PlotNo = '" & .txtPlotNo.Text & "' AND OwenedByClient = 'N';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                    Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!MastNo))
                    If Not IsNull(rsLIST!TypeOfMast) Then
                        MyList.SubItems(1) = CStr(rsLIST!TypeOfMast)
                    End If
                    
                    If Not IsNull(rsLIST!MediaSize) Then
                            MyList.SubItems(2) = CStr(rsLIST!MediaSize)
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
Public Sub LeasedMasts()
On Error GoTo err
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Structure No", .ListView3.Width / 3 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , " Media", .ListView3.Width / 3
                .ListView3.ColumnHeaders.Add , , "Size", .ListView3.Width / 3
 
                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlotMast Where ContractNo = '" & .txtContractNo.Text & "' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!MastNo))
                        If Not IsNull(rsLIST!TypeOfMast) Then
                            MyList.SubItems(1) = CStr(rsLIST!TypeOfMast)
                        End If
                        
                        If Not IsNull(rsLIST!MediaSize) Then
                                MyList.SubItems(2) = CStr(rsLIST!MediaSize)
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

Public Sub getContractNo()
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM ODASMLeaseAgreement WHERE PlotNo = '" & frmODASMAllocation.txtPlotNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsFindRecord.RecordCount = 0 Then Exit Sub
    frmODASMAllocation.txtContractNo.Text = rsFindRecord!ContractNo
    Set rsFindRecord = Nothing
End Sub

Public Sub getLANDLORDS()
On Error GoTo err
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Land Lord No", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Names", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 3
 
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                If bsearchRECORD = True Then
                        strSQL = "SELECT * FROM ODASPAccount Where CompanyName like '%" & Trim(.txtSearchName.Text) & "%' and Status = 'A' AND AccountType = 'LLORD' Order by AccountNo;"
                Else
                        strSQL = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' oRDER BY AccountNo;"
                End If
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!AccountNo))
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                                MyList.SubItems(2) = CStr(rsLIST!Status)
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
Public Sub showALLLandlords()
On Error GoTo err
        With frmODASMSiteRegistration
        
                .ListALLLandLords.ListItems.Clear
                .ListALLLandLords.ColumnHeaders.Clear
                
                .ListALLLandLords.ColumnHeaders.Add , , "Land Lord No", .ListALLLandLords.Width / 7 ', lvwColumnCenter
                .ListALLLandLords.ColumnHeaders.Add , , "Names", .ListALLLandLords.Width / 1.8
                .ListALLLandLords.ColumnHeaders.Add , , "Status", .ListALLLandLords.Width / 7
 
                .ListALLLandLords.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                If bsearchRECORD = True Then
                        strSQL = "SELECT * FROM ODASPAccount Where (CompanyName like '%" & Trim(.txtLandLordName.Text) & "%' OR PhysicalAddress like '%" & Trim(.txtLandLordName.Text) & "%' OR PostalAddress like '%" & Trim(.txtLandLordName.Text) & "%' and MobileNo like '%" & Trim(.txtLandLordName) & "%' ) and Status = 'A' AND AccountType = 'LLORD' Order by AccountNo;"
                Else
                        strSQL = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' oRDER BY AccountNo;"
                End If
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListALLLandLords.ListItems.Add(, , CStr(rsLIST!AccountNo))
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                                MyList.SubItems(2) = CStr(rsLIST!Status)
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

Public Sub getLANDLORDTYPE()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Type", .ListView1.Width / 2 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2
 
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPAccountType WHERE AccountType = 'LLORD';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                DF = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!AccountType))
                            If Not IsNull(rsLIST!AccountTypeDescription) Then
                                MyList.SubItems(1) = CStr(rsLIST!AccountTypeDescription)
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

Public Sub showALLFreeSites()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.JobBriefNo is Null and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Sites"
                While Not rsLIST.EOF
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                        End If
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                        End If
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(3) = CStr(rsLIST!PlotName)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                            MyList.SubItems(4) = CStr(rsLIST!Status)
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
Public Sub ListALLSitesToFree()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.JobBriefNo is Not Null and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Sites"
                While Not rsLIST.EOF
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                        End If
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                        End If
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(3) = CStr(rsLIST!PlotName)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                            MyList.SubItems(4) = CStr(rsLIST!Status)
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
Public Sub showALLSitesReserved()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Client", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot,ODASPPlotSite,ODASMJobBrief,ODASPAccount where ODASPPlotSite.Status ='SITE-RESERVED' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASMJobBrief.AccountNo = ODASPAccount.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Sites"
                While Not rsLIST.EOF
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                        End If
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                        End If
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(3) = CStr(rsLIST!PlotName)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(4) = CStr(rsLIST!CompanyName)
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

Public Sub showALLSitesAllocated()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Client", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Date Started", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot,ODASPPlotSite,ODASMJobBrief,ODASPAccount where ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASMJobBrief.AccountNo = ODASPAccount.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Sites"
                While Not rsLIST.EOF
                            
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                            
                            If Not IsNull(rsLIST!SiteDetails) Then
                                MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                            End If
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                            End If
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(3) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(4) = CStr(rsLIST!CompanyName)
                            End If
                            
                            If Not IsNull(rsLIST!JCStartDate) Then
                                MyList.SubItems(5) = Format(rsLIST!JCStartDate, "dd/mm/yyyy")
                            End If
                            
                            If Not IsNull(rsLIST!JCExpiryDate) Then
                                MyList.SubItems(6) = Format(rsLIST!JCExpiryDate, "dd/mm/yyyy")
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

Public Sub showALLSitesToFree()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Product", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Date Started", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot,ODASPPlotSite,ODASMJobBrief where ODASPPlotSite.Status ='SITE-ALLOCATED' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASPPlotSite.JobBriefNo is not null and ODASPPlotSite.JCExpiryDate <'" & Format(Date, "MMMM dd,yyyy") & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Sites"
                While Not rsLIST.EOF
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                        End If
                        
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(2) = CStr(rsLIST!PlotName)
                        End If
                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(3) = CStr(rsLIST!ProductCode)
                        End If
                        
                                                   
                        If Not IsNull(rsLIST!JCStartDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!JCStartDate)
                        End If
                        
                        If Not IsNull(rsLIST!JCExpiryDate) Then
                            MyList.SubItems(5) = CStr(rsLIST!JCExpiryDate)
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

Public Sub AllSitesOnRoadReserve()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "BillBoard No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Details", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot,ODASPPlotMast where ODASPPlot.OnRoadReserve ='Y' and ODASPPlot.PlotNo = ODASPPlotMast.PlotNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Sites"
                While Not rsLIST.EOF
                            
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                    
                    If Not IsNull(rsLIST!MastDetails) Then
                        MyList.SubItems(1) = CStr(rsLIST!MastDetails)
                    End If
                    
                    If Not IsNull(rsLIST!PlotNo) Then
                        MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                    End If
                    If Not IsNull(rsLIST!PlotName) Then
                        MyList.SubItems(3) = CStr(rsLIST!PlotName)
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

Public Sub AllNonEagleStructures()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "BillBoard No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Details", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot,ODASPPlotMast where ODASPPlotMast.OwenedByClient ='Y' and ODASPPlot.PlotNo = ODASPPlotMast.PlotNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Sites"
                While Not rsLIST.EOF
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                        
                        If Not IsNull(rsLIST!MastDetails) Then
                            MyList.SubItems(1) = CStr(rsLIST!MastDetails)
                        End If
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                        End If
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(3) = CStr(rsLIST!PlotName)
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

Public Sub ShowSiteSchedule()
On Error GoTo err
    
        With frmODASMSiteSchedule
                
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 2 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Schedule Date", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "JobBrief Item", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Reserved", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Allocated", .ListView1.Width / 2

                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
            
                strSQL = "SELECT * FROM ODASMSiteSchedule where SiteNo = '" & CurrentRecord & "' and ScheduleDate >= '" & Format(Date, "MMMM dd,yyyy") & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsLIST.RecordCount = 0 Then Exit Sub
                
                SchedulingMain.ProgressBar1.Visible = True
                SchedulingMain.ProgressBar1.Value = 0
                SchedulingMain.ProgressBar1.Max = rsLIST.RecordCount
                SchedulingMain.ProgressBar1.Min = 0
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                            
                            If Not IsNull(rsLIST!ScheduleDAte) Then
                                MyList.SubItems(1) = CStr(rsLIST!ScheduleDAte)
                            End If
                            
                            If Not IsNull(rsLIST!JobBriefItemNo) Then
                                    MyList.SubItems(2) = CStr(rsLIST!JobBriefItemNo)
                            End If
                            
                            If Not IsNull(rsLIST!Reserved) Then
                                    MyList.SubItems(3) = CStr(rsLIST!Reserved)
                            End If
                            
                            If Not IsNull(rsLIST!Allocated) Then
                                    MyList.SubItems(4) = CStr(rsLIST!Allocated)
                            End If

                        rsLIST.MoveNext
                        SchedulingMain.ProgressBar1.Value = SchedulingMain.ProgressBar1.Value + 1
                Wend
                Set MyList = Nothing
        End With
SchedulingMain.ProgressBar1.Visible = False
Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub ShowBSSchedule()
On Error GoTo err
    
        With frmODASMSiteSchedule
                
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 2 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Schedule Date", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "JobBrief Item", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Reserved", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Allocated", .ListView1.Width / 2

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
            
                strSQL = "SELECT * FROM ODASMSiteSchedule where SiteNo = '" & CurrentRecord & "' and Reserved = 'Y' and ScheduleDate >= '" & Format(Date, "MMMM dd,yyyy") & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsLIST.RecordCount = 0 Then Exit Sub
                
                SchedulingMain.ProgressBar1.Visible = True
                SchedulingMain.ProgressBar1.Value = 0
                SchedulingMain.ProgressBar1.Max = rsLIST.RecordCount
                SchedulingMain.ProgressBar1.Min = 0
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                            
                            If Not IsNull(rsLIST!ScheduleDAte) Then
                                MyList.SubItems(1) = CStr(rsLIST!ScheduleDAte)
                            End If
                            
                            If Not IsNull(rsLIST!JobBriefItemNo) Then
                                    MyList.SubItems(2) = CStr(rsLIST!JobBriefItemNo)
                            End If
                            
                            If Not IsNull(rsLIST!Reserved) Then
                                    MyList.SubItems(3) = CStr(rsLIST!Reserved)
                            End If
                            
                            If Not IsNull(rsLIST!Allocated) Then
                                    MyList.SubItems(4) = CStr(rsLIST!Allocated)
                            End If

                        rsLIST.MoveNext
                        SchedulingMain.ProgressBar1.Value = SchedulingMain.ProgressBar1.Value + 1
                Wend
                Set MyList = Nothing
        End With
SchedulingMain.ProgressBar1.Visible = False
Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub ShowBBSchedule()
On Error GoTo err
    
        With frmODASMSiteSchedule
                
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Billboard No", .ListView1.Width / 2 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Schedule Date", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "JobBrief Item", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Reserved", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Allocated", .ListView1.Width / 2.5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
            
                strSQL = "SELECT * FROM ODASMBillBoardSchedule where MastNo = '" & CurrentRecord & "' and Reserved ='Y' and ScheduleDate >= '" & Format(Date, "MMMM dd,yyyy") & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsLIST.RecordCount = 0 Then Exit Sub
                
                SchedulingMain.ProgressBar1.Visible = True
                SchedulingMain.ProgressBar1.Value = 0
                SchedulingMain.ProgressBar1.Max = rsLIST.RecordCount
                SchedulingMain.ProgressBar1.Min = 0
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                            
                            If Not IsNull(rsLIST!ScheduleDAte) Then
                                MyList.SubItems(1) = CStr(rsLIST!ScheduleDAte)
                            End If
                            
                            If Not IsNull(rsLIST!JobBriefItemNo) Then
                                    MyList.SubItems(2) = CStr(rsLIST!JobBriefItemNo)
                            End If
                            
                            If Not IsNull(rsLIST!Reserved) Then
                                    MyList.SubItems(3) = CStr(rsLIST!Reserved)
                            End If
                            
                            If Not IsNull(rsLIST!Allocated) Then
                                    MyList.SubItems(4) = CStr(rsLIST!Allocated)
                            End If

                        rsLIST.MoveNext
                        SchedulingMain.ProgressBar1.Value = SchedulingMain.ProgressBar1.Value + 1
                Wend
                Set MyList = Nothing
        End With
SchedulingMain.ProgressBar1.Visible = False
Exit Sub

err:
        If err.Number = 3265 Then Resume Next
        ErrorMessage
End Sub
Public Sub RateSchedules()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Location", .ListView1.Width / 3.5
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Advert", .ListView1.Width / 4.5
                .ListView1.ColumnHeaders.Add , , "Rates Payable", .ListView1.Width / 5.5
                .ListView1.ColumnHeaders.Add , , "Rates DueDate", .ListView1.Width / 5
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot, ODASMJobBriefItems, ODASPPlotSite, ODASPTown, ODASMJobBrief,ODASMCouncilRatesPayable WHERE ODASPPlot.PlotNo = ODASPPlotSite.PLotNo and ODASMJobBriefItems.JobBriefItemNo = ODASPPlotSite.JobBriefItemNo and ODASMJobBriefItems.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASPPlot.Towncode = ODASPTown.towncode and ODASPPlotSite.SiteNo = ODASMCouncilRatesPayable.SiteNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                While Not rsLIST.EOF
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(1) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!Town) Then
                            MyList.SubItems(2) = CStr(rsLIST!Town)
                        End If
                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(3) = CStr(rsLIST!ProductCode)
                        End If
                        If Not IsNull(rsLIST!RatePayable) Then
                            MyList.SubItems(4) = CStr(rsLIST!RatePayable)
                        End If
                        If Not IsNull(rsLIST!RateDueDate) Then
                            MyList.SubItems(5) = CStr(rsLIST!RateDueDate)
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

Public Sub AllSiteSchedule()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 3.5
                .ListView1.ColumnHeaders.Add , , "LRNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 4.5
                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot where status='SITE-ACQUIRED'"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                While Not rsLIST.EOF
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                        
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotName)
                        End If
                        
                        If Not IsNull(rsLIST!LRNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!LRNo)
                        End If
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(3) = CStr(rsLIST!PhysicalLocation)
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
Public Sub AllPlotRents()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Location", .ListView1.Width / 2.5
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Rent", .ListView1.Width / 6
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot, ODASPTown WHERE ODASPPlot.AnnualRent is not null and ODASPPlot.Towncode = ODASPTown.towncode;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Sites"
                While Not rsLIST.EOF
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                        
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(1) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!Town) Then
                            MyList.SubItems(2) = CStr(rsLIST!Town)
                        End If
                        If Not IsNull(rsLIST!AnnualRent) Then
                            MyList.SubItems(3) = FormatNumber(rsLIST!AnnualRent, 2, vbUseDefault, vbUseDefault, vbUseDefault)
                        End If
                        
                        rsLIST.MoveNext
                Wend
                .ListView1.ColumnHeaders(4).Alignment = lvwColumnRight

                
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage

End Sub

Public Sub showALLSitesUnAllocated()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.Status ='SITE-AVAILABLE' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Sites"
                While Not rsLIST.EOF
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                        End If
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                        End If
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(3) = CStr(rsLIST!PlotName)
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

Public Sub showALLSitesWithoutProperties()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "BBoard No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Faces", .ListView1.Width / 15
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Physcical Location", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Structure", .ListView1.Width / 7
                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot,ODASPPlotMast where ODASPPlotMast.PlotNo = ODASPPlot.PlotNo and (ODASPPlotMast.PropertiesAssigned='N' or ODASPPlotMast.PropertiesAssigned is null);"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Billboards"
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                        
                        If Not IsNull(rsLIST!NoofSites) Then
                            MyList.SubItems(1) = CStr(rsLIST!NoofSites)
                        End If
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                        End If
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(3) = CStr(rsLIST!PlotName)
                        End If
                        
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(4) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!TypeOfMast) Then
                            MyList.SubItems(5) = CStr(rsLIST!TypeOfMast)
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

Public Sub showALLSitesWithoutPropertiesbyNumber()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "BBoard No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Faces", .ListView1.Width / 15
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Physcical Location", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Structure", .ListView1.Width / 7
                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot,ODASPPlotMast where ODASPPlotMast.PlotNo = ODASPPlot.PlotNo and MastNo like '%" & CurrentRecord & "%';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Billboards"
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                            
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                    
                    If Not IsNull(rsLIST!NoofSites) Then
                        MyList.SubItems(1) = CStr(rsLIST!NoofSites)
                    End If
                    
                    If Not IsNull(rsLIST!PlotNo) Then
                        MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                    End If
                    If Not IsNull(rsLIST!PlotName) Then
                        MyList.SubItems(3) = CStr(rsLIST!PlotName)
                    End If
                    
                    If Not IsNull(rsLIST!PhysicalLocation) Then
                        MyList.SubItems(4) = CStr(rsLIST!PhysicalLocation)
                    End If
                    
                    If Not IsNull(rsLIST!TypeOfMast) Then
                        MyList.SubItems(5) = CStr(rsLIST!TypeOfMast)
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

Public Sub showALLSitesbyNumber()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Face No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 8 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Physcical Location", .ListView1.Width / 3
                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot,ODASPPlotSite where ODASPPlotSite.PlotNo = ODASPPlot.PlotNo and SiteNo like '%" & CurrentRecord & "%';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Billboard faces"
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                            
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(2) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(3) = CStr(rsLIST!PhysicalLocation)
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

Public Sub showALLSitesWithoutPropertiesbyTown()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "BBoard No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Faces", .ListView1.Width / 15
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Physcical Location", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Structure", .ListView1.Width / 7
                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot,ODASPPlotMast,ODASPTown where ODASPPLot.TownCode = ODASPTown.TownCode and ODASPPlotMast.PlotNo = ODASPPlot.PlotNo and Town like '" & CurrentRecord & "%';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Billboards"
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                            
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                            
                            If Not IsNull(rsLIST!NoofSites) Then
                                MyList.SubItems(1) = CStr(rsLIST!NoofSites)
                            End If
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                            End If
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(3) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(4) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!TypeOfMast) Then
                                MyList.SubItems(5) = CStr(rsLIST!TypeOfMast)
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

Public Sub showALLSitesWithoutPropertiesbyPlotName()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "BBoard No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Faces", .ListView1.Width / 15
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Physcical Location", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Structure", .ListView1.Width / 7
                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot,ODASPPlotMast where ODASPPlotMast.PlotNo = ODASPPlot.PlotNo and PlotName like '%" & CurrentRecord & "%';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Billboards"
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                            
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                            
                            If Not IsNull(rsLIST!NoofSites) Then
                                MyList.SubItems(1) = CStr(rsLIST!NoofSites)
                            End If
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                            End If
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(3) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(4) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!TypeOfMast) Then
                                MyList.SubItems(5) = CStr(rsLIST!TypeOfMast)
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

Public Sub showALLSitesWithoutPropertiesbyLandlordName()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "BBoard No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Faces", .ListView1.Width / 15
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Physcical Location", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Structure", .ListView1.Width / 7
                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot,ODASPPlotMast,ODASPAccount where ODASPPlot.AccountNo = ODASPAccount.AccountNo and ODASPPlotMast.PlotNo = ODASPPlot.PlotNo and CompanyName like '%" & CurrentRecord & "%';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Billboards"
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                            
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                            
                            If Not IsNull(rsLIST!NoofSites) Then
                                MyList.SubItems(1) = CStr(rsLIST!NoofSites)
                            End If
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                            End If
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(3) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(4) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!TypeOfMast) Then
                                MyList.SubItems(5) = CStr(rsLIST!TypeOfMast)
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

Public Sub setALLAcquiredSites()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotName", .ListView1.Width / 3
                 .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot  where (OnRoadReserve = 'N' or OnRoadReserve is null) ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                DF = rsLIST.RecordCount
                
                rsLIST.MoveFirst
                Do While rsLIST.EOF <> True
                
                Set rsFindRecord = New ADODB.Recordset
                rsFindRecord.Open "SELECT *  FROM ODASPPlotMast where OwenedByClient = 'N' and  (LeasePrepared = 'N' or LeasePrepared is null) and PlotNo = '" & rsLIST!PlotNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsFindRecord.RecordCount > 0 Then
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                            
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                            DoEvents
                    End If
                rsLIST.MoveNext
            Loop
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub
Public Sub setALLAcquiredSitesForRenewal()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotName", .ListView1.Width / 3
                 .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot  where (OnRoadReserve = 'N' or OnRoadReserve is null) ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                DF = rsLIST.RecordCount
                
                rsLIST.MoveFirst
                Do While rsLIST.EOF <> True
                
                Set rsFindRecord = New ADODB.Recordset
                rsFindRecord.Open "SELECT *  FROM ODASPPlotMast where OwenedByClient = 'N' and PlotNo = '" & rsLIST!PlotNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsFindRecord.RecordCount > 0 Then
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                            
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                    End If
                rsLIST.MoveNext
            Loop
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub getALLAllocatedSites()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotName", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot P, ODASMLeaseAgreement LA where P.PlotNo = LA.PlotNo  and (LA.Approved = 'N' or LA.Approved is null) ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If

                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(2) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(3) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(4) = CStr(rsLIST!AccountNo)
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

Public Sub getALLsites()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 1

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlotSite;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                            
                            If Not IsNull(rsLIST!SiteDetails) Then
                                MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
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
Public Sub getALLApprovedMasts()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Mast No", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 1

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlotSite;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!MastNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!MastNo)
                            End If
                            
                            If Not IsNull(rsLIST!SiteDetails) Then
                                MyList.SubItems(3) = CStr(rsLIST!SiteDetails)
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

Public Sub getALLApprovedSites()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotName", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot P, ODASMLeaseAgreement LA where  P.PlotNo = LA.PlotNo and (LA.Authorized is null or LA.Authorized = 'N') and LA.Approved = 'Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If

                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(2) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(3) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(3) = CStr(rsLIST!AccountNo)
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

Public Sub getALLsitesRatesPrepared()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 2.5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Job Item No on Site", .ListView1.Width / 2 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Current Year", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT  distinct CR.CurrentYear, CR.SiteNo,CR.JobBriefItemNo  FROM  ODASMCouncilRateDue CR ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Leases"
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!CurrentYear) Then
                            MyList.SubItems(2) = CStr(rsLIST!CurrentYear)
                        End If
                        
                        If Not IsNull(rsLIST!JobBriefItemNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!JobBriefItemNo)
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

Public Sub getOneALLsitesRatesPrepared()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "JobBriefNo", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Current Year", .ListView1.Width / 1.5 ', lvwColumnCenter
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT  distinct CR.CurrentYear, CR.SiteNo,CR.JobBriefItemNo  FROM  ODASMCouncilRateDue CR Where CR.SiteNo = '" & CurrentRecord & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Leases"
                While Not rsLIST.EOF
                                                 
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                    
                    If Not IsNull(rsLIST!JobBriefItemNo) Then
                        MyList.SubItems(1) = CStr(rsLIST!JobBriefItemNo)
                    End If
                    
                    If Not IsNull(rsLIST!CurrentYear) Then
                        MyList.SubItems(2) = CStr(rsLIST!CurrentYear)
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

Public Sub getCURRENTLEASES()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot P,ODASPAccount A, ODASMLeaseAgreement LA where  LA.PlotNo = P.PlotNo AND P.AccountNo = A.AccountNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Leases"
                While Not rsLIST.EOF
                                                 
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                    
                    If Not IsNull(rsLIST!PlotNo) Then
                        MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                    End If
                    
                    If Not IsNull(rsLIST!CommencementDate) Then
                        MyList.SubItems(2) = Format(rsLIST!CommencementDate, "dd/mm/yyyy")
                    End If
                    
                    If Not IsNull(rsLIST!expirydate) Then
                        MyList.SubItems(3) = Format(rsLIST!expirydate, "dd/mm/yyyy")
                    End If
                    
                    If Not IsNull(rsLIST!PhysicalLocation) Then
                        MyList.SubItems(4) = CStr(rsLIST!PhysicalLocation)
                    End If
                    If Not IsNull(rsLIST!CompanyName) Then
                        MyList.SubItems(5) = CStr(rsLIST!CompanyName)
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

Public Sub getOneCURRENTLEASE()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot P,ODASPAccount A, ODASMLeaseAgreement LA where  LA.PlotNo = P.PlotNo AND P.AccountNo = A.AccountNo and LA.ContractNo like '%" & CurrentRecord & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Leases"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!CommencementDate) Then
                                MyList.SubItems(2) = Format(rsLIST!CommencementDate, "dd/mm/yyyy")
                            End If
                            
                            If Not IsNull(rsLIST!expirydate) Then
                                MyList.SubItems(3) = Format(rsLIST!expirydate, "dd/mm/yyyy")
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(4) = CStr(rsLIST!PhysicalLocation)
                            End If
                            If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(5) = CStr(rsLIST!CompanyName)
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

Public Sub getLEASESTerminatedLandLord()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Termination Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Termination By", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Reasons", .ListView1.Width / 3
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASPPlot, ODASMLeaseAgreement,ODASPTerminationReasons where ODASMLeaseAgreement.TerminationDate >= '" & Format(frmODASSearchSitesTerminated.txtStartDate.Text, "MMMM dd,yyyy") & "' and ODASMLeaseAgreement.TerminationDate <= '" & Format(frmODASSearchSitesTerminated.txtLastDate.Text, "MMMM dd,yyyy") & "' and ODASMLeaseAgreement.Terminated = 'Y' and ODASPPlot.PlotNo = ODASMLeaseAgreement.PlotNo and ODASMLeaseAgreement.TerminationCode = ODASPTerminationReasons.TerminationCode and  ODASMLeaseAgreement.TerminationCode='01';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Leases"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!NoticeDate) Then
                                MyList.SubItems(3) = Format(rsLIST!NoticeDate, "dd/mm/yyyy")
                            End If
                             
                            If Not IsNull(rsLIST!TerminationDate) Then
                                MyList.SubItems(4) = Format(rsLIST!TerminationDate, "dd/mm/yyyy")
                            End If
                            If Not IsNull(rsLIST!TerminationReason) Then
                                MyList.SubItems(5) = CStr(rsLIST!TerminationReason)
                            End If
                            If Not IsNull(rsLIST!Narration) Then
                                MyList.SubItems(6) = CStr(rsLIST!Narration)
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
Public Sub getLEASESTerminatedCompany()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Termination Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Termination By", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Reasons", .ListView1.Width / 3
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASPPlot, ODASMLeaseAgreement,ODASPTerminationReasons where ODASMLeaseAgreement.TerminationDate >= '" & Format(frmODASSearchSitesTerminated.txtStartDate.Text, "MMMM dd,yyyy") & "' and ODASMLeaseAgreement.TerminationDate <= '" & Format(frmODASSearchSitesTerminated.txtLastDate.Text, "MMMM dd,yyyy") & "' and ODASMLeaseAgreement.Terminated = 'Y' and ODASPPlot.PlotNo = ODASMLeaseAgreement.PlotNo and ODASMLeaseAgreement.TerminationCode = ODASPTerminationReasons.TerminationCode and  ODASMLeaseAgreement.TerminationCode='02';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Leases"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!NoticeDate) Then
                                MyList.SubItems(3) = Format(rsLIST!NoticeDate, "dd/mm/yyyy")
                            End If
                             
                            If Not IsNull(rsLIST!TerminationDate) Then
                                MyList.SubItems(4) = Format(rsLIST!TerminationDate, "dd/mm/yyyy")
                            End If
                            If Not IsNull(rsLIST!TerminationReason) Then
                                MyList.SubItems(5) = CStr(rsLIST!TerminationReason)
                            End If
                            If Not IsNull(rsLIST!Narration) Then
                                MyList.SubItems(6) = CStr(rsLIST!Narration)
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

Public Sub getLEASESDUEToEXPIRE()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotName", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Contract", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Commencent Date", .ListView1.Width / 5
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMLeaseAgreement L,ODASPPlot P WHERE (L.Terminated is null or L.Terminated ='N') AND P.PlotNo = L.PlotNo and P.ExpiryDate > '" & Format(Date, "MMMM dd,yyyy") & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Leases"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(3) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!ContractNo) Then
                                MyList.SubItems(4) = CStr(rsLIST!ContractNo)
                            End If
                            If Not IsNull(rsLIST!expirydate) Then
                                MyList.SubItems(5) = Format(rsLIST!expirydate, "dd/mm/yyyy")
                            End If
                            If Not IsNull(rsLIST!CommencementDate) Then
                                MyList.SubItems(6) = Format(rsLIST!CommencementDate, "dd/mm/yyyy")
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

Public Sub getLEASESRentNotPaid()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 8 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotName", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Contract", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Rent", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Rent Due", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 8

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot where  rentDueDate <= '" & Format(Date, "yyyy/mm/dd") & "' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!LandLordNo) Then
                                MyList.SubItems(3) = CStr(rsLIST!LandLordNo)
                            End If
                            
                            If Not IsNull(rsLIST!ContractNo) Then
                                MyList.SubItems(4) = CStr(rsLIST!ContractNo)
                            End If
                            
                            If Not IsNull(rsLIST!AnnualRent) Then
                                MyList.SubItems(5) = CStr(rsLIST!AnnualRent)
                            End If
                            
                            If Not IsNull(rsLIST!RentDue) Then
                                MyList.SubItems(6) = CStr(rsLIST!RentDue)
                            End If

                            If Not IsNull(rsLIST!RentDueDate) Then
                                MyList.SubItems(7) = CStr(rsLIST!RentDueDate)
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

Public Sub loadDEFAULTSRates()
On Error GoTo err
    
        With frmODASMCouncilRates
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Media code", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Media Size", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Payment Mode", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPLandRate where TownCode = '" & .txtTownCode & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MediaCode))
                            
                            If Not IsNull(rsLIST!MediaSize) Then
                                MyList.SubItems(1) = CStr(rsLIST!MediaSize)
                            End If
                            If Not IsNull(rsLIST!PaymentMode) Then
                                MyList.SubItems(2) = CStr(rsLIST!PaymentMode)
                            End If
                            If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(3) = CStr(rsLIST!Amount)
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

Public Sub getSITESRatesNotSet()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "PlotName", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "JobBrief No", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlotSite, ODASPPlot where  ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.Rates is null;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                While Not rsLIST.EOF
                                                 
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                    
                    If Not IsNull(rsLIST!TownCode) Then
                        MyList.SubItems(1) = CStr(rsLIST!TownCode)
                    End If
                    If Not IsNull(rsLIST!PlotName) Then
                        MyList.SubItems(2) = CStr(rsLIST!PlotName)
                    End If
                    If Not IsNull(rsLIST!PhysicalLocation) Then
                        MyList.SubItems(3) = CStr(rsLIST!PhysicalLocation)
                    End If
                  
                    If Not IsNull(rsLIST!SiteDetails) Then
                        MyList.SubItems(4) = CStr(rsLIST!SiteDetails)
                    End If
                    If Not IsNull(rsLIST!JobBriefNo) Then
                        MyList.SubItems(5) = CStr(rsLIST!JobBriefNo)
                    End If
                    If Not IsNull(rsLIST!Status) Then
                        MyList.SubItems(6) = CStr(rsLIST!Status)
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

Public Sub getSITESRateRentPaid()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 8 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotName", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Contract", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Rent Due", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Rate Due", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 8

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlotSite, ODASPPlot where  ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotSite.rateDueDate > '" & Format(Date, "yyyy/mm/dd") & "' and ODASPPlot.rentDueDate > '" & Format(Date, "yyyy/mm/dd") & "' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!LandLordNo) Then
                                MyList.SubItems(3) = CStr(rsLIST!LandLordNo)
                            End If
                            
                            If Not IsNull(rsLIST!ContractNo) Then
                                MyList.SubItems(4) = CStr(rsLIST!ContractNo)
                            End If
                            
                            If Not IsNull(rsLIST!RentDue) Then
                                MyList.SubItems(5) = CStr(rsLIST!RentDue)
                            End If
                            
                            If Not IsNull(rsLIST!RateDue) Then
                                MyList.SubItems(6) = CStr(rsLIST!RateDue)
                            End If

                            If Not IsNull(rsLIST!RateDueDate) Then
                                MyList.SubItems(7) = CStr(rsLIST!RateDueDate)
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

Public Sub getPLOTRentNotPaid()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 8 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotName", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Contract", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Annual Rent", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Rent Paid", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Rent Due", .ListView1.Width / 8

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot, ODASPPlotMast where  (ODASPPlot.RentPaid =0 or ODASPPlot.RentPaid < ODASPPlot.AnnualRent) and ODASPPlot.AnnualRent>0 and ODASPPlot.PlotNo= ODASPPlotMast.PLotNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                Dim MyList As ListItem, RentDue
                
                       While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(3) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!ContractNo) Then
                                MyList.SubItems(4) = CStr(rsLIST!ContractNo)
                            End If
                            
                            If Not IsNull(rsLIST!AnnualRent) Then
                                MyList.SubItems(5) = CStr(rsLIST!AnnualRent)
                            End If
                            If Not IsNull(rsLIST!RentPaid) Then
                                MyList.SubItems(6) = CStr(rsLIST!RentPaid)
                            End If
                            RentDue = rsLIST!AnnualRent - rsLIST!RentPaid
                            MyList.SubItems(7) = RentDue
                            
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
        ErrorMessage

End Sub

Public Sub getPLOTRentRateDues()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 8 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotName", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Annual Rent", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Rent Paid", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Rent Due", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Annual Rate", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Rate Paid", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Rate Due", .ListView1.Width / 8

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot where  (ODASPPlot.RentPaid = 0 or ODASPPlot.RentPaid < ODASPPlot.AnnualRent) and ODASPPlot.AnnualRent > 0 AND (ODASPPlot.RatePaid =0 or ODASPPlot.RatePaid < ODASPPlot.AnnualRate);"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                Dim MyList As ListItem, RentDue As Variant, RateDue As Variant
                       While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!AnnualRent) Then
                                MyList.SubItems(3) = CStr(rsLIST!AnnualRent)
                            End If
                            If Not IsNull(rsLIST!RentPaid) Then
                                MyList.SubItems(4) = CStr(rsLIST!RentPaid)
                            End If
                            RentDue = rsLIST!AnnualRent - rsLIST!RentPaid
                            MyList.SubItems(5) = RentDue
                             
                            If Not IsNull(rsLIST!AnnualRate) Then
                                MyList.SubItems(6) = CStr(rsLIST!AnnualRate)
                            End If
                            If Not IsNull(rsLIST!RatePaid) Then
                                MyList.SubItems(7) = CStr(rsLIST!RatePaid)
                            End If
                            RateDue = rsLIST!AnnualRate - rsLIST!RatePaid
                            MyList.SubItems(8) = RateDue

                            
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage

End Sub

Public Sub getSITESRateNotPaid()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Annual Rate", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Rate Due", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT (ODASPPlotSite.Rates)as Rate,(ODASPPlotSite.RateDue)as RateD,(ODASPPlotSite.RateDueDate) as DueDate, ODASPPlotSite.*,ODASPPlot.*  FROM ODASPPlotSite, ODASPPlot where  ODASPPlotSite.PlotNo = ODASPPlot.PlotNo and ODASPPlotSite.rateDueDate <= '" & Format(Date, "yyyy/mm/dd") & "' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                While Not rsLIST.EOF
                                                 
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
                    
                    If Not IsNull(rsLIST!SiteDetails) Then
                        MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                    End If
                    
                    If Not IsNull(rsLIST!PhysicalLocation) Then
                        MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                    End If
                    
                    If Not IsNull(rsLIST!rate) Then
                        MyList.SubItems(3) = CStr(rsLIST!rate)
                    End If
                    
                    If Not IsNull(rsLIST!RateD) Then
                        MyList.SubItems(4) = CStr(rsLIST!RateD)
                    End If

                    If Not IsNull(rsLIST!DueDate) Then
                        MyList.SubItems(5) = CStr(rsLIST!DueDate)
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

Public Sub getALLContracts()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT (L.ExpiryDate) as EDates, L.*, P.*  FROM ODASMLeaseAgreement L, ODASPPlot P where L.Assigned = 'Y' and (L.Terminated = 'N' or L.Terminated is null) AND P.PlotNo = L.PlotNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsLIST.RecordCount = 0 Then Exit Sub
                rsLIST.MoveFirst
                While Not rsLIST.EOF
                
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(2) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(3) = CStr(rsLIST!AccountNo)
                            End If
                            If Not IsNull(rsLIST!EDates) Then
                                MyList.SubItems(4) = Format(rsLIST!EDates, "dd/mm/yyyy")
                            End If
Continue:
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With
 
Exit Sub

err:
    If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub getALLNACADAContracts()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT (L.ExpiryDate) as EDates, L.*, P.*  FROM ODASMLeaseAgreement L, ODASPPlot P where L.Assigned = 'Y' and (L.Terminated = 'N' or L.Terminated is null) AND P.PlotNo = L.PlotNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsLIST.RecordCount = 0 Then Exit Sub
                rsLIST.MoveFirst
                While Not rsLIST.EOF
                
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            If Not IsNull(rsLIST!EDates) Then
                                MyList.SubItems(3) = Format(rsLIST!EDates, "dd/mm/yyyy")
                            End If
Continue:
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With
 
Exit Sub

err:
    If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub getOneContract()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT (L.ExpiryDate) as EDates, L.*, P.*  FROM ODASMLeaseAgreement L, ODASPPlot P where L.Assigned = 'Y' and (L.Terminated = 'N' or L.Terminated is null) AND P.PlotNo = L.PlotNo and L.ContractNo like '%" & CurrentRecord & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsLIST.RecordCount = 0 Then Exit Sub
                rsLIST.MoveFirst
                While Not rsLIST.EOF
                
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            If Not IsNull(rsLIST!EDates) Then
                                MyList.SubItems(3) = Format(rsLIST!EDates, "dd/mm/yyyy")
                            End If
Continue:
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub getNoticesPrepared()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticePrepared = 'Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Notices"
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                        End If
                        
                        If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If
                        
                        If Not IsNull(rsLIST!NoticeDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!NoticeDate)
                        End If
                         
                        If Not IsNull(rsLIST!NoticePreparedBy) Then
                            MyList.SubItems(4) = CStr(rsLIST!NoticePreparedBy)
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

Public Sub getALLNoticesAuthorized()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Authorization Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Authorized By", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticeAuthorized = 'Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Notices"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!NoticeAuthorizationDate) Then
                                MyList.SubItems(3) = CStr(rsLIST!NoticeAuthorizationDate)
                            End If
                             
                            If Not IsNull(rsLIST!AuthorizedBy) Then
                                MyList.SubItems(4) = CStr(rsLIST!AuthorizedBy)
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

Public Sub getNOTICESAPPROVED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Approved By", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Approval date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticeApproved = 'Y' AND (NoticeAUTHORIZED = 'N' OR NoticeAUTHORIZED is null);"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Notices"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!NoticeApprovedBy) Then
                                MyList.SubItems(3) = CStr(rsLIST!NoticeApprovedBy)
                            End If
                             
                            If Not IsNull(rsLIST!NoticeApprovalDate) Then
                                MyList.SubItems(4) = CStr(rsLIST!NoticeApprovalDate)
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

Public Sub getNOTICESAUTHORIZED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticeAuthorized = 'Y' and (NoticeDispatched is null or NoticeDispatched = 'N') ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Notices"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!NoticeDate) Then
                                MyList.SubItems(3) = CStr(rsLIST!NoticeDate)
                            End If
                             
                            If Not IsNull(rsLIST!NoticePreparedBy) Then
                                MyList.SubItems(4) = CStr(rsLIST!NoticePreparedBy)
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

Public Sub getAllNoticesSent()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Notices"
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                        End If
                        
                        If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If
                        
                        If Not IsNull(rsLIST!NoticeDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!NoticeDate)
                        End If
                         
                        If Not IsNull(rsLIST!NoticePreparedBy) Then
                            MyList.SubItems(4) = CStr(rsLIST!NoticePreparedBy)
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

Public Sub getNoticesSent()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y' and (NoticeReceived is null or NoticeReceived = 'N');"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                        End If
                        
                        If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If
                        
                        If Not IsNull(rsLIST!NoticeDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!NoticeDate)
                        End If
                         
                        If Not IsNull(rsLIST!NoticePreparedBy) Then
                            MyList.SubItems(4) = CStr(rsLIST!NoticePreparedBy)
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

Public Sub getNoticesReceived()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y' and NoticeReceived = 'Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                        End If
                        
                        If Not IsNull(rsLIST!LandLordNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!LandLordNo)
                        End If
                        
                        If Not IsNull(rsLIST!NoticeDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!NoticeDate)
                        End If
                         
                        If Not IsNull(rsLIST!NoticePreparedBy) Then
                            MyList.SubItems(4) = CStr(rsLIST!NoticePreparedBy)
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

Public Sub showNoticesApproved()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Item No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Client", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT (JBI.ExpiryDate)as EDate,JBI.*,JB.*,A.*  FROM ODASMJobBriefItems JBI,ODASPAccount A,ODASMJobBrief JB where JBI.NoticeReceived = 'Y' and JBI.NoticeApproved = 'Y' and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Contracts"
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!SiteDetails) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteDetails)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!EDate) Then
                            MyList.SubItems(3) = Format(rsLIST!EDate, "dd/mm/yyyy")
                        End If
                         
                        If Not IsNull(rsLIST!NoticeReceivedDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!NoticeReceivedDate)
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

Public Sub showRenewalNoticesApproved()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Item No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Client", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT (JBI.ExpiryDate)as EDate,JBI.*,JB.*,A.*  FROM ODASMJobBriefItems JBI,ODASPAccount A,ODASMJobBrief JB where JBI.NoticeAuthorized = 'N' and JBI.NoticeApproved = 'Y' and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Contracts"
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefItemNo))
                        
                        If Not IsNull(rsLIST!SiteNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteNo)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!EDate) Then
                            MyList.SubItems(3) = Format(rsLIST!EDate, "dd/mm/yyyy")
                        End If
                         
                        If Not IsNull(rsLIST!NoticeReceivedDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!NoticeReceivedDate)
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

Public Sub showNoticesAuthorized()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Item No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Client", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Renewal period", .ListView1.Width / 5
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT (JBI.ExpiryDate)as EDate,JBI.*,JB.*,A.*  FROM ODASMJobBriefItems JBI,ODASPAccount A,ODASMJobBrief JB where JBI.NoticeAuthorized = 'Y' and JBI.Status = 'NOTICE-AUTHORIZED' and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Contracts"
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefItemNo))
                        
                        If Not IsNull(rsLIST!SiteNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteNo)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!EDate) Then
                            MyList.SubItems(3) = Format(rsLIST!EDate, "dd/mm/yyyy")
                        End If
                         
                        If Not IsNull(rsLIST!NoticeReceivedDate) Then
                            MyList.SubItems(4) = Format(rsLIST!NoticeReceivedDate, "dd/mm/yyyy")
                        End If
                        If Not IsNull(rsLIST!RenewalPeriod) Then
                            MyList.SubItems(5) = CStr(rsLIST!RenewalPeriod)
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

Public Sub showNoticesReceived()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Item No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Client", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT (JBI.ExpiryDate)as EDate,JBI.*,JB.*,A.*  FROM ODASMJobBriefItems JBI,ODASPAccount A,ODASMJobBrief JB where JBI.NoticeReceived = 'Y' and (JBI.NoticeApproved = 'N' or JBI.NoticeApproved is null) and JBI.JobBriefNo = JB.JobBriefNo and JB.AccountNo = A.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Contracts"
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefItemNo))
                        
                        If Not IsNull(rsLIST!SiteNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!SiteNo)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!EDate) Then
                            MyList.SubItems(3) = Format(rsLIST!EDate, "dd/mm/yyyy")
                        End If
                         
                        If Not IsNull(rsLIST!NoticeReceivedDate) Then
                            MyList.SubItems(4) = Format(rsLIST!NoticeReceivedDate, "dd/mm/yyyy")
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

Public Sub getCONTRACTSToRenew()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement A,ODASPPlot P where A.NoticeDispatched = 'Y' and (A.Renewed is null or A.Renewed = 'N') and A.ReasonsForNotice = 'Renewal of Contract' and A.AccountNo = P.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Contracts"
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                        End If
                        
                        If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If
                        
                        If Not IsNull(rsLIST!expirydate) Then
                            MyList.SubItems(3) = Format(rsLIST!expirydate, "dd/mm/yyyy")
                        End If
                         
                        If Not IsNull(rsLIST!NoticeDate) Then
                            MyList.SubItems(4) = Format(rsLIST!NoticeDate, "dd/mm/yyyy")
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

Public Sub getCONTRACTSRenewed()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Date Renewed", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "New Expiry Date", .ListView1.Width / 4
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseRenewals LR, ODASMLeaseAgreement A,ODASPPlot P where  LR.ContractNo = A.ContractNo and A.AccountNo = P.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Contracts"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!expirydate) Then
                                MyList.SubItems(3) = Format(rsLIST!expirydate, "dd/mm/yyyy")
                            End If
                             
                            If Not IsNull(rsLIST!RenewalDate) Then
                                MyList.SubItems(4) = Format(rsLIST!RenewalDate, "dd/mm/yyyy")
                            End If
                            If Not IsNull(rsLIST!expirydate) Then
                                MyList.SubItems(5) = Format(rsLIST!expirydate, "dd/mm/yyyy")
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

Public Sub getCONTRACTSNotices()
On Error GoTo err
    
        With frmODASMPrepareNotice
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Reason for Notice", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Date Prepared", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 4
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement A,ODASPPlot P where A.NoticePrepared = 'Y' and  A.AccountNo = P.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Contracts"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!ReasonsForNotice) Then
                                MyList.SubItems(3) = CStr(rsLIST!ReasonsForNotice)
                            End If
                             
                            If Not IsNull(rsLIST!NoticeDate) Then
                                MyList.SubItems(4) = CStr(rsLIST!NoticeDate)
                            End If
                            If Not IsNull(rsLIST!NoticePreparedBy) Then
                                MyList.SubItems(5) = CStr(rsLIST!NoticePreparedBy)
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

Public Sub getCONTRACTSTerminated()
On Error GoTo err
    
        With frmODASMLeaseAgreement
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Termination Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Termination By", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Reasons", .ListView1.Width / 4
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPTerminationReasons R,ODASMLeaseAgreement A,ODASPPlot P where A.Terminated = 'Y' and  A.AccountNo = P.AccountNo and R.TerminationCode = A.TerminationCode;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Contracts"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!TerminationDate) Then
                                MyList.SubItems(3) = CStr(rsLIST!TerminationDate)
                            End If
                             
                            If Not IsNull(rsLIST!TerminatedBy) Then
                                MyList.SubItems(4) = CStr(rsLIST!TerminatedBy)
                            End If
                            If Not IsNull(rsLIST!TerminationReason) Then
                                MyList.SubItems(5) = CStr(rsLIST!TerminationReason)
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

Public Sub getCONTRACTSToTerminate()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y' and (Terminated is null or Terminated = 'N') and ReasonsForNotice = 'Termination of Contract';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Contracts"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                            End If
                            
                            If Not IsNull(rsLIST!NoticeDate) Then
                                MyList.SubItems(3) = CStr(rsLIST!NoticeDate)
                            End If
                             
                            If Not IsNull(rsLIST!NoticePreparedBy) Then
                                MyList.SubItems(4) = CStr(rsLIST!NoticePreparedBy)
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

Public Sub NoticesPrepared()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticePrepared = 'Y' and (NoticeApproved = 'N' or NoticeApproved is null);"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Notices"
                While Not rsLIST.EOF
                                                 
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                    
                    If Not IsNull(rsLIST!PlotNo) Then
                        MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                    End If
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                        MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!NoticeDate) Then
                        MyList.SubItems(3) = CStr(rsLIST!NoticeDate)
                    End If
                     
                    If Not IsNull(rsLIST!NoticePreparedBy) Then
                        MyList.SubItems(4) = CStr(rsLIST!NoticePreparedBy)
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

Public Sub RenewalNoticesReceived()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job ItemNo", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticePrepared = 'Y' and (NoticeApproved = 'N' or NoticeApproved is null);"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Notices"
                While Not rsLIST.EOF
                                                 
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                    
                    If Not IsNull(rsLIST!PlotNo) Then
                        MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                    End If
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                        MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!NoticeDate) Then
                        MyList.SubItems(3) = CStr(rsLIST!NoticeDate)
                    End If
                     
                    If Not IsNull(rsLIST!NoticePreparedBy) Then
                        MyList.SubItems(4) = CStr(rsLIST!NoticePreparedBy)
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

Public Sub getNOTICESDISPATCHED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Notice Date", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASMLeaseAgreement where NoticeDispatched = 'Y' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Notices"
                While Not rsLIST.EOF
                                                 
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                    
                    If Not IsNull(rsLIST!PlotNo) Then
                        MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                    End If
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                        MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!NoticeDate) Then
                        MyList.SubItems(3) = CStr(rsLIST!NoticeDate)
                    End If
                     
                    If Not IsNull(rsLIST!NoticePreparedBy) Then
                        MyList.SubItems(4) = CStr(rsLIST!NoticePreparedBy)
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

Public Sub showALLTerminationReasons()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Termination Code", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Reason", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Company", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Others", .ListView1.Width / 5
                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPTerminationReasons where Status = 'A ';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                                                 
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!TerminationCode))
                        
                        If Not IsNull(rsLIST!TerminationReason) Then
                            MyList.SubItems(1) = CStr(rsLIST!TerminationReason)
                        End If
                        
                        If Not IsNull(rsLIST!Company) Then
                            MyList.SubItems(2) = CStr(rsLIST!Company)
                        End If
                        
                        If Not IsNull(rsLIST!LandLord) Then
                            MyList.SubItems(3) = CStr(rsLIST!LandLord)
                        End If
                        
                        If Not IsNull(rsLIST!Others) Then
                            MyList.SubItems(4) = CStr(rsLIST!Others)
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

Public Sub showALLPlotsToExpire()
On Error GoTo err
    
        With frmODASSitesToExpire
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Company Name", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "DOC", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                If frmODASSitesToExpire.strReport = "" Then
                        strSQL = "SELECT ODASPPlot.*,ODASPAccount.COmpanyName  FROM ODASPPlot, ODASPAccount,ODASMLeaseAgreement WHERE ODASMLeaseAgreement.PlotNo=ODASPPlot.PlotNo AND ODASPPLot.ExpiryDate >= '" & Format(frmODASSitesToExpire.txtStartDate.Text, "yyyy/mm/dd") & "' and ODASPPLot.ExpiryDate <= '" & Format(frmODASSitesToExpire.txtLastDate.Text, "yyyy/mm/dd") & "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo and (ODASMLeaseAgreement.terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL);"
                Else
                        strSQL = "SELECT ODASPPlot.*,ODASPAccount.COmpanyName  FROM ODASPPlot, ODASPAccount,ODASMLeaseAgreement WHERE  ODASMLeaseAgreement.PlotNo=ODASPPlot.PlotNo AND ODASPPLot.ExpiryDate <= '" & Format(frmODASSitesToExpire.txtLastDate.Text, "yyyy/mm/dd") & "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo  AND (ODASMLeaseAgreement.terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL);"
                End If
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                            
                            If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!CommencementDate) Then
                                MyList.SubItems(3) = Format(rsLIST!CommencementDate, "dd/mm/yyyy")
                            End If
                            
                            If Not IsNull(rsLIST!expirydate) Then
                                MyList.SubItems(4) = Format(rsLIST!expirydate, "dd/mm/yyyy")
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
Public Sub getALLNewSites()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "BillBoard No", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Details", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "DOC", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                         
                strSQL = "SELECT *  FROM ODASPPlot, ODASPPLotMast, ODASPAccount where ODASPPLotMast.LeasePrepared = 'Y' and ODASPPLotMast.CommencementDate >= '" & Format(frmODASSearchSiteNewSites.txtStartDate.Text, "yyyy/mm/dd") & "' and ODASPPLotMast.CommencementDate <= '" & Format(frmODASSearchSiteNewSites.txtLastDate.Text, "yyyy/mm/dd") & "' and ODASPPlot.PlotNo = ODASPPLotMast.PLotNo and ODASPAccount.AccountNo=ODASPPlot.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                            
                            If Not IsNull(rsLIST!MastDetails) Then
                                MyList.SubItems(1) = CStr(rsLIST!MastDetails)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!CommencementDate) Then
                                MyList.SubItems(3) = Format(rsLIST!CommencementDate, "dd/mm/yyyy")
                            End If
                            
                            If Not IsNull(rsLIST!expirydate) Then
                                MyList.SubItems(4) = Format(rsLIST!expirydate, "dd/mm/yyyy")
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
'Public Sub GetRentNotRequisitioned()
'On Error GoTo err
'
'        With Screen.ActiveForm
'
'            .ListView1.ListItems.Clear
'            .ListView1.ColumnHeaders.Clear
'
'            .ListView1.ColumnHeaders.Add , , "Installment ", .ListView1.Width / 6
'            .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 6
'            .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6
'            .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
'            .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 6
'            .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
'
'            .ListView1.View = lvwReport
'
'            Dim rsLIST As ADODB.Recordset
'            Set rsLIST = New ADODB.Recordset
'
'            strSQL = "Select * from ODASMInstallment I, ODASPAccount A Where I.PaymentDueDate <= '" & Format(frmODASSearchSitesNotPaid.txtLastDate.Text, "yyyy/mm/dd") & "' and (I.Requisitioned = 'N' or I.Requisitioned is null) and I.AccountNo = A.AccountNo;"
'            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
'
'            If rsLIST.RecordCount <> 0 Then
'
'
'            Dim MyList As ListItem
'
'            While Not rsLIST.EOF
'
'                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
'
'                    If Not IsNull(rsLIST!ContractNo) Then
'                            MyList.SubItems(1) = CStr(rsLIST!ContractNo)
'                    End If
'
'                    If Not IsNull(rsLIST!PaymentDueDate) Then
'                            MyList.SubItems(2) = CStr(rsLIST!PaymentDueDate)
'                    End If
'
'                    If Not IsNull(rsLIST!PaymentDue) Then
'                            MyList.SubItems(3) = CStr(rsLIST!PaymentDue)
'                    End If
'
'                    If Not IsNull(rsLIST!AccountNo) Then
'                            MyList.SubItems(4) = CStr(rsLIST!AccountNo)
'                    End If
'
'                    If Not IsNull(rsLIST!CompanyName) Then
'                            MyList.SubItems(5) = CStr(rsLIST!CompanyName)
'                    End If
'
'                    rsLIST.MoveNext
'            Wend
'            End If
'            Set MyList = Nothing
'        End With
'
'Exit Sub
'
'err:
'If err.Number = 3265 Then Resume Next
'ErrorMessage
'End Sub
Public Sub GetInstallmentToEdit()
On Error GoTo err
    
        With Screen.ActiveForm
        
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Installment ", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
        
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "Select * from ODASMInstallment I, ODASPAccount A Where I.ContractNo= '" & frmODASPEditIncrement.txtContractNo.Text & "' and (I.Requisitioned = 'N' or I.Requisitioned is null) and I.AccountNo = A.AccountNo;"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsLIST.RecordCount <> 0 Then
            
            
            Dim MyList As ListItem
                       
            While Not rsLIST.EOF
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                    
                    If Not IsNull(rsLIST!ContractYear) Then
                            MyList.SubItems(1) = CStr(rsLIST!ContractYear)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!PaymentDueDate)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDue) Then
                            MyList.SubItems(3) = CStr(rsLIST!PaymentDue)
                    End If
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(4) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(5) = CStr(rsLIST!CompanyName)
                    End If
                    
                    rsLIST.MoveNext
            Wend
            End If
            Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetRentRequisitioned()
On Error GoTo err
    
        With Screen.ActiveForm
        
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Installment ", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
             .ListView1.ColumnHeaders.Add , , "Requisition Date", .ListView1.Width / 6
             .ListView1.ColumnHeaders.Add , , "Serial", .ListView1.Width / 6
        
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "Select * from ODASMInstallment I, ODASPAccount A,ODASMLeaseAgreement L Where (L.PlotNo=I.PlotNo AND L.ContractNo=I.ContractNo AND (L.Terminated='N' OR L.Terminated IS NULL) ) AND I.Requisitioned = 'Y' and I.AccountNo = A.AccountNo AND (I.VoucherDate>='" & Format(.DTPStartDate, "yyyy/MM/dd") & "' AND I.VoucherDate<='" & Format(.DTPLastDate, "yyyy/MM/dd") & "') AND I.ChequeNo IS NULL;"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsLIST.RecordCount <> 0 Then
            
            .ProgressBar1.Visible = True
            .ProgressBar1.Value = 0: .ProgressBar1.Min = 0: .ProgressBar1.Max = rsLIST.RecordCount
            
            Dim MyList As ListItem
                       
            While Not rsLIST.EOF
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                    
                    If Not IsNull(rsLIST!ContractNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!ContractNo)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!PaymentDueDate)
                    End If
        
                    If Not IsNull(rsLIST!AmountPaid) Then
                            MyList.SubItems(3) = CStr(rsLIST!AmountPaid)
                    End If
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(4) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(5) = CStr(rsLIST!CompanyName)
                    End If
                    
                    If Not IsNull(rsLIST!VoucherDate) Then
                            MyList.SubItems(6) = CStr(rsLIST!VoucherDate)
                    End If
                    
                    If Not IsNull(rsLIST!InstallmentNo) Then
                            MyList.SubItems(7) = CStr(rsLIST!InstallmentNo)
                    End If
                    
                    .ProgressBar1.Value = .ProgressBar1.Value + 1
                    rsLIST.MoveNext
            Wend
            .ProgressBar1.Visible = False
            End If
            Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetVouchersPrepared()
On Error GoTo err
    
        With Screen.ActiveForm
                    strSQL = "SELECT     V.VoucherNo, V.VoucherDate, I.ContractNo, I.Installment, PA.CompanyName, V.Amount,V.Printed FROM ODASMVoucher AS V INNER JOIN ODASPAccount AS PA ON V.AccountNo = PA.AccountNo  INNER JOIN  ODASMInstallment AS I ON V.VoucherNo = I.VoucherNo WHERE (v.voucherdate>='" & Format(.DTPStartDate.Value, "yyyy/MM/dd") & "' AND v.voucherdate<='" & Format(.DTPLastDate.Value, "yyyy/MM/dd") & "')"
                    FillList strSQL, .ListView1
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub


Public Sub getALLSitesExpired()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "BillBoard No", .ListView1.Width / 8 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "DOC", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 8

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT ODASPPlot.PlotNo, ODASPAccount.CompanyName, ODASPPlot.*  FROM ODASPPlot, ODASPAccount where ODASPPlot.AccountNo = ODASPPlot.AccountNo and ODASPPlot.ExpiryDate <= '" & Format(Date, "yyyy/mm/dd") & "' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                            
                            If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                            
                            If Not IsNull(rsLIST!CommencementDate) Then
                                MyList.SubItems(3) = Format(rsLIST!CommencementDate, "dd/mm/yyyy")
                            End If
                            
                            If Not IsNull(rsLIST!expirydate) Then
                                MyList.SubItems(4) = Format(rsLIST!expirydate, "dd/mm/yyyy")
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

Public Sub showALLTOWNS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Town Code", .ListView1.Width / 2 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 2

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT TownCode, Town FROM ODASPTown ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!TownCode))
                        If Not IsNull(rsLIST!Town) Then
                            MyList.SubItems(1) = CStr(rsLIST!Town)
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
Public Sub showALLINSTALLMENTSPAID()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Rent", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Payment Date ", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Flag", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT TownCode, Town FROM ODASMInstallment I Where I.CurrentPeriod = '" & .txtCurrentPeriod & "' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                    
                        If Not IsNull(rsLIST!Total) Then
                            MyList.SubItems(1) = CStr(rsLIST!Town)
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

Public Sub showALLINSTALLMENTSDUE()
On Error GoTo err
    
        With frmODASMSiteRegistration
        
                .ListALLInstallments.ListItems.Clear
                .ListALLInstallments.ColumnHeaders.Clear
                
                .ListALLInstallments.ColumnHeaders.Add , , "Installment No", .ListALLInstallments.Width / 8 ', lvwColumnCenter
                .ListALLInstallments.ColumnHeaders.Add , , "Rent", .ListALLInstallments.Width / 8
                .ListALLInstallments.ColumnHeaders.Add , , "Payment Date ", .ListALLInstallments.Width / 8
                .ListALLInstallments.ColumnHeaders.Add , , "Flag", .ListALLInstallments.Width / 8
                .ListALLInstallments.ColumnHeaders.Add , , "Current Year", .ListALLInstallments.Width / 8 ', lvwColumnCenter
                .ListALLInstallments.ColumnHeaders.Add , , "Payment Period", .ListALLInstallments.Width / 8
                .ListALLInstallments.ColumnHeaders.Add , , "Payment Mode ", .ListALLInstallments.Width / 8
                .ListALLInstallments.ColumnHeaders.Add , , "Inv No", .ListALLInstallments.Width / 8
                .ListALLInstallments.ColumnHeaders.Add , , "Payment Due", .ListALLInstallments.Width / 8

                .ListALLInstallments.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMInstallment I Where I.PlotNo='" & .txtPlotNo & "' AND I.ContractNo = '" & .txtContractNo & "' Order by InstallmentNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                        Set MyList = .ListALLInstallments.ListItems.Add(, , CStr(rsLIST!InstallmentNo))
                    
                        If Not IsNull(rsLIST!TotalRent) Then
                            MyList.SubItems(1) = FormatNumber(rsLIST!TotalRent, 2, vbUseDefault, vbUseDefault, vbUseDefault)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!PaymentDueDate)
                        End If

                        If Not IsNull(rsLIST!PaymentFlag) Then
                            MyList.SubItems(3) = CStr(rsLIST!PaymentFlag)
                        End If
                        
                        If Not IsNull(rsLIST!ContractYear) Then
                            MyList.SubItems(4) = CStr(rsLIST!ContractYear)
                        End If
                        
                        If Not IsNull(rsLIST!CurrentPeriod) Then
                            MyList.SubItems(5) = CStr(rsLIST!CurrentPeriod)
                        End If

                        If Not IsNull(rsLIST!PaymentMode) Then
                            MyList.SubItems(6) = CStr(rsLIST!PaymentMode)
                        End If
                        
                        If Not IsNull(rsLIST!InvoiceNo) Then
                            MyList.SubItems(7) = CStr(rsLIST!InvoiceNo)
                        End If
                        If Not IsNull(rsLIST!PaymentDue) Then
                            MyList.SubItems(8) = FormatNumber(rsLIST!PaymentDue, 2, vbUseDefault, vbUseDefault, vbUseDefault)
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

Public Sub showALLRentPAIDThisPeriod()
On Error GoTo err
    
        With frmODASSearchPaidSites
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Mast No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Location", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Starting", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Ending", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Installment", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Payment Date ", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Amount Paid", .ListView1.Width / 9

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where ODASMInstallment.PaymentDate >= '" & Format(frmODASSearchPaidSites.txtStartDate.Text, "yyyy/MM/dd") & "' and (ODASMInstallment.PaymentFlag = 'Y' or ODASMInstallment.PaymentFlag = 'P') AND ODASMInstallment.PaymentDate <= '" & Format(frmODASSearchPaidSites.txtLastDate.Text, "yyyy/MM/dd") & "' and ODASPPlot.PlotNo = ODASMInstallment.ContractNo and ODASPPlot.AccountNo = ODASPAccount.AccountNo Order by ODASMInstallment.PaymentDate;"

                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                    
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(1) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!CommencementDate) Then
                            MyList.SubItems(3) = Format(rsLIST!CommencementDate, "dd/mm/yyyy")
                        End If
                        
                        If Not IsNull(rsLIST!expirydate) Then
                            MyList.SubItems(4) = Format(rsLIST!expirydate, "dd/mm/yyyy")
                        End If
                            
                        If Not IsNull(rsLIST!Installment) Then
                            MyList.SubItems(5) = CStr(rsLIST!Installment)
                        End If
                            
                        If Not IsNull(rsLIST!ContractNo) Then
                            MyList.SubItems(6) = CStr(rsLIST!ContractNo)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentDate) Then
                            MyList.SubItems(7) = CStr(rsLIST!PaymentDate)
                        End If
                        
                        If Not IsNull(rsLIST!AmountPaid) Then
                            MyList.SubItems(8) = FormatNumber(rsLIST!AmountPaid, 2, vbUseDefault, vbUseDefault, vbUseDefault)
                        End If
                        

                        
                     rsLIST.MoveNext
                Wend
                
                .ListView1.ColumnHeaders(9).Alignment = lvwColumnRight

                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub
Public Sub showALLRentPendingPayment()
       strSQL = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where (ODASMInstallment.PaymentDueDate >= '" & Format(Screen.ActiveForm.txtStartDate.Value, "yyyy/mm/dd") & "' AND ODASMInstallment.PaymentDueDate <= '" & Format(Screen.ActiveForm.txtLastDate.Value, "yyyy/mm/dd") & "') and ODASMInstallment.PaymentFlag = 'N' and ODASMInstallment.Balance > 0 and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo ORDER BY CompanyName;"
        FillList strSQL, frmODASSearchSitesNotPaid.ListView1
End Sub

Public Sub showALLRentPendingPaymentAsAtASingleDate()
       strSQL = "SELECT ODASMInstallment.* FROM ODASMInstallment,ODASPPlot,ODASPAccount,ODASMLeaseAgreement Where ODASMLeaseAgreement.PlotNo=ODASMInstallment.PlotNo AND ODASMLeaseAgreement.ContractNo=ODASMInstallment.ContractNo AND (ODASMLeaseAgreement.Terminated='N' OR ODASMLeaseAgreement.Terminated IS NULL) AND (ODASMInstallment.PaymentDueDate <= '" & Format(Screen.ActiveForm.txtLastDate.Value, "yyyy/mm/dd") & "') and ODASMInstallment.PaymentFlag = 'N' and ODASMInstallment.Balance > 0 and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo ORDER BY CompanyName;"
        Debug.Print strSQL
        FillList strSQL, frmODASSearchSitesNotPaid.ListView1

End Sub

Public Sub showALLRentVouchersPrepared()
       strSQL = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where ODASMInstallment.DateRequisitioned <= '" & Format(Screen.ActiveForm.txtLastDate.Value, "yyyy/mm/dd") & "' and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.Requisitioned ='Y')  ORDER BY CompanyName;"
        FillList strSQL, frmODASSearchSitesNotPaid.ListView1
End Sub

Public Sub showALLRentPendingConfirmation()
        strSQL = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where (ODASMInstallment.DateRequisitioned >= '" & Format(Screen.ActiveForm.txtLastDate.Value, "yyyy/mm/dd") & "' AND ODASMInstallment.DateRequisitioned <= '" & Format(Screen.ActiveForm.txtLastDate.Value, "yyyy/mm/dd") & "') and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.Requisitioned ='Y') AND (ODASMInstallment.PaymentFlag='N' OR ODASMInstallment.PaymentFlag IS NULL)   ORDER BY CompanyName;"
        FillList strSQL, frmODASSearchSitesNotPaid.ListView1

End Sub

Public Sub showALLRentWithPaymentsConfirmed()
       strSQL = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where (ODASMInstallment.DateRequisitioned >= '" & Format(Screen.ActiveForm.txtLastDate.Value, "yyyy/mm/dd") & "' AND ODASMInstallment.DateRequisitioned <= '" & Format(Screen.ActiveForm.txtLastDate.Value, "yyyy/mm/dd") & "') and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo AND (ODASMInstallment.PaymentFlag ='Y')  ORDER BY CompanyName;"
       Debug.Print strSQL
        FillList strSQL, frmODASSearchSitesNotPaid.ListView1

End Sub

Public Sub showALLRentNOTPAIDThisPeriod()
On Error GoTo err
    
        With frmODASSearchSitesNotPaid
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Mast No", .ListView1.Width / 9 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Location", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Starting", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Ending", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Sides", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Payment Due Date", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "AmountDue ", .ListView1.Width / 9

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPAccount Where (ODASMInstallment.PaymentDueDate >= '" & Format(Screen.ActiveForm.txtStartDate.Value, "yyyy/mm/dd") & "' AND ODASMInstallment.PaymentDueDate <= '" & Format(Screen.ActiveForm.txtLastDate.Value, "yyyy/mm/dd") & "') and ODASMInstallment.PaymentFlag = 'N' and ODASMInstallment.Balance > 0 and ODASPPlot.PlotNo=ODASMInstallment.PlotNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo ;"
Debug.Print strSQL
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                    
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(1) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!CommencementDate) Then
                            MyList.SubItems(3) = Format(rsLIST!CommencementDate, "dd/mm/yyyy")
                        End If
                        
                        If Not IsNull(rsLIST!expirydate) Then
                            MyList.SubItems(4) = Format(rsLIST!expirydate, "dd/mm/yyyy")
                        End If
                            
                        If Not IsNull(rsLIST!NoofSites) Then
                            MyList.SubItems(5) = FormatNumber(rsLIST!NoofSites, 2, vbUseDefault, vbUseDefault, vbUseDefault)
                        End If
                            
                        If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(6) = CStr(rsLIST!PaymentDueDate)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentDue) Then
                            MyList.SubItems(7) = FormatNumber(rsLIST!PaymentDue, 2, vbUseDefault, vbUseDefault, vbUseDefault)
                        End If
                     rsLIST.MoveNext
                Wend
                .ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
                .ListView1.ColumnHeaders(8).Alignment = lvwColumnRight

                
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showALLRentDueThisPeriod()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Rent", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Balance", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMInstallment I Where I.CurrentPeriod >= '" & frmODASYearsearch.cbostartperiod & "' and PaymentFlag = 'N' AND I.CurrentPeriod <= '" & frmODASYearsearch.cboendperiod & "'   ;"
               
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                    
                        If Not IsNull(rsLIST!TotalRent) Then
                            MyList.SubItems(1) = CStr(rsLIST!TotalRent)
                        End If
                                                
                        If Not IsNull(rsLIST!Balance) Then
                            MyList.SubItems(2) = CStr(rsLIST!Balance)
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

Public Sub showTOWNS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "Town Code", .ListView2.Width / 2 ', lvwColumnCenter
                .ListView2.ColumnHeaders.Add , , "Town", .ListView2.Width / 2

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT TownCode, Town FROM ODASPTown ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!TownCode))
                        If Not IsNull(rsLIST!Town) Then
                            MyList.SubItems(1) = CStr(rsLIST!Town)
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


Public Sub showALLACTIVESITES()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 2 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Site Name", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Location", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 2

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlot Where Status = 'ACTIVE';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
                        
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotName)
                        End If
                        
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!TownCode) Then
                            MyList.SubItems(3) = CStr(rsLIST!TownCode)
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

Public Sub showALLSITES()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Site No", .ListView3.Width / 4 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , "Mast No", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Plot No", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Details", .ListView3.Width / 4

                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlotSite, ODASpPlot Where ODASPPlotSite.PlotNo = ODASPPlot.PlotNo AND (ODASPPlotSite.MastNo = '" & .txtMastNo.Text & "' or ODASPPlotSite.PlotNo = '" & .txtPlotNo.Text & "');"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!SiteNo))
                        
                        If Not IsNull(rsLIST!MastNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!MastNo)
                        End If
                        
                        If Not IsNull(rsLIST!PlotNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!PlotNo)
                        End If
                        
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(3) = CStr(rsLIST!PlotName)
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
Public Sub showALLSITESByLandlord()
On Error GoTo err
    
        With frmODASPLandLord
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "Plot No", .ListView2.Width / 7 ', lvwColumnCenter
                .ListView2.ColumnHeaders.Add , , "Location", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "DOC", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Expiry", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Rent Due", .ListView2.Width / 7

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASPAccount, ODASpPlot Where ODASPAccount.AccountNo = ODASPPlot.AccountNo and ODASPPlot.AccountNo = '" & .txtLandLordNo.Text & "' Order by ODASPPlot.PlotNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!PlotNo))
                       
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(1) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!CommencementDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!CommencementDate)
                        End If
                        
                        If Not IsNull(rsLIST!expirydate) Then
                            MyList.SubItems(3) = CStr(rsLIST!expirydate)
                        End If
                        
                        If Not IsNull(rsLIST!AnnualRent) Then
                            MyList.SubItems(4) = FormatNumber(rsLIST!AnnualRent, 2, vbUseDefault, vbUseDefault, vbUseDefault)
                        End If
                        
                        
                        rsLIST.MoveNext
                Wend
                
                .ListView2.ColumnHeaders(5).Alignment = lvwColumnRight

                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showALLContracts()
On Error GoTo errMSG
With frmODASMSiteRegistration
    strSQL = "SELECT ContractNo,PlotNo,CommencementDate,ExpiryDate,LeaseDuration,AnnualRent FROM ODASMLeaseAgreement WHERE PLotNo LIKE '" & .txtPlotNo & "' ORDER BY CommencementDate DESC"
    FillList strSQL, .ListView4
End With
Exit Sub
errMSG:
    ErrorMessage
End Sub

Public Sub showALLUNALLOCATEDPLOTS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "BillBoard No", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Location", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "No Of Faces", .ListView1.Width / 4
                

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlotMast, ODASPPlot Where ODASPPlotMast.PlotNo= ODASPPlot.PlotNo and (LeasePrepared = 'N' or LeasePrepared is null) and OwenedByClient ='N' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                    
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotName)
                        End If
                        
                        If Not IsNull(rsLIST!MastDetails) Then
                            MyList.SubItems(2) = CStr(rsLIST!MastDetails)
                        End If
                        
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(3) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!NoofSites) Then
                            MyList.SubItems(4) = CStr(rsLIST!NoofSites)
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

Public Sub getALLOPERATIONS()
On Error GoTo err

    With Screen.ActiveForm
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Operation", .ListView1.Width / 2#
            .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2
            
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASPOperationType"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListView1.View = lvwList
                Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
            
            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!OperationType))
            
                If Not IsNull(rsLIST!Description) Then
                    MyList.SubItems(1) = CStr(rsLIST!Description)
                End If
                 
                rsLIST.MoveNext
                
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub


Public Sub showALLDEPARTMENTS()
On Error GoTo err

    With Screen.ActiveForm
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Department", .ListView1.Width / 2#
            .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2
            
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASPDepartment where Status = 'A'"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListView1.View = lvwList
                Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
            
            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!DepartmentCode))
            
            
                If Not IsNull(rsLIST!DepartmentDescription) Then
                    MyList.SubItems(1) = CStr(rsLIST!DepartmentDescription)
                End If
                 
                rsLIST.MoveNext
                
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub showALLBillBoardSchedule()
On Error GoTo err

    With Screen.ActiveForm
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 4#
            .ListView1.ColumnHeaders.Add , , "Billboard", .ListView1.Width / 4
            .ListView1.ColumnHeaders.Add , , "CommencementDate", .ListView1.Width / 4
            .ListView1.ColumnHeaders.Add , , "ExpiryDate", .ListView1.Width / 4
            
            
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASppLOT where BillBoard = 'Y';" ' and ExpiryDate= > Date"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListView1.View = lvwList
                Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
            
            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
            
            
                If Not IsNull(rsLIST!BillBoard) Then
                    MyList.SubItems(1) = CStr(rsLIST!BillBoard)
                End If
                If Not IsNull(rsLIST!CommencementDate) Then
                    MyList.SubItems(2) = CStr(rsLIST!CommencementDate)
                End If
                If Not IsNull(rsLIST!expirydate) Then
                    MyList.SubItems(3) = CStr(rsLIST!expirydate)
                End If
                
                 
                rsLIST.MoveNext
                
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub showALLDEPTS()
On Error GoTo err

    With Screen.ActiveForm
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Department", .ListView1.Width / 2#
            .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2
            
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASPDepartment where Status = 'A'"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListView1.View = lvwList
                Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
                    Set rsCONTROL = New ADODB.Recordset
            
                    strCONTROL = "SELECT * FROM ODASPMediaTask where DepartmentCode = '" & rsLIST!DepartmentCode & "' and MediaCode = '" & frmODASPMediaTask.txtMediaCode.Text & "' "
                    rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    If rsCONTROL.EOF Or rsCONTROL.BOF Then
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!DepartmentCode))
                        
                            If Not IsNull(rsLIST!DepartmentDescription) Then
                                MyList.SubItems(1) = CStr(rsLIST!DepartmentDescription)
                            End If
                    End If
                    
                    rsLIST.MoveNext
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub


Public Sub showACTUALPROPERTIES()
  On Error GoTo err

    With Screen.ActiveForm
            .ListView2.ListItems.Clear
            .ListView2.ColumnHeaders.Clear
            
            .ListView2.ColumnHeaders.Add , , "Property Code", .ListView2.Width / 5
            .ListView2.ColumnHeaders.Add , , "Description", .ListView2.Width / 1

            .ListView2.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASMSiteProperties,ODASPProperties where sITENo = '" & frmODASPAssignProperties.txtSiteNo & "' and ODASMSiteProperties.PropertyCode=ODASPProperties.PropertyCode; "
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListView2.View = lvwList
                Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
            
            Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!PropertyCode))
                       
                If Not IsNull(rsLIST!PropertyDescription) Then
                    MyList.SubItems(1) = CStr(rsLIST!PropertyDescription)
                End If
                rsLIST.MoveNext
                
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub showACTUALPROPERTIES1()
  On Error GoTo err

    With Screen.ActiveForm
            .ListActualProperties.ListItems.Clear
            .ListActualProperties.ColumnHeaders.Clear
            
            .ListActualProperties.ColumnHeaders.Add , , "Property Code", .ListActualProperties.Width / 5
            .ListActualProperties.ColumnHeaders.Add , , "Description", .ListActualProperties.Width / 1

            .ListActualProperties.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASMSiteProperties,ODASPProperties where SiteNo = '" & .txtSiteNo & "' and ODASMSiteProperties.PropertyCode=ODASPProperties.PropertyCode; "
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListActualProperties.View = lvwList
                Set MyList = .ListActualProperties.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
            
            Set MyList = .ListActualProperties.ListItems.Add(, , CStr(rsLIST!PropertyCode))
                       
                If Not IsNull(rsLIST!PropertyDescription) Then
                    MyList.SubItems(1) = CStr(rsLIST!PropertyDescription)
                End If
                rsLIST.MoveNext
                
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub


Public Sub showALLLandLORDSites()
On Error GoTo err

    With Screen.ActiveForm
            .ListView2.ListItems.Clear
            .ListView2.ColumnHeaders.Clear
            
            .ListView2.ColumnHeaders.Add , , "Contract", .ListView2.Width / 6
            .ListView2.ColumnHeaders.Add , , "Plot", .ListView2.Width / 6
            .ListView2.ColumnHeaders.Add , , "LandLord", .ListView2.Width / 6
            .ListView2.ColumnHeaders.Add , , "Signed", .ListView2.Width / 6
            .ListView2.ColumnHeaders.Add , , "Agreement Date", .ListView2.Width / 6
            .ListView2.ColumnHeaders.Add , , "Signed By", .ListView2.Width / 6

            .ListView2.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASMLeaseAgreement where AccountNo = '" & Screen.ActiveForm.txtLandLordNo & "' "
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListView2.View = lvwList
                Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
            
            Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!ContractNo))
                
                If Not IsNull(rsLIST!ContractNo) Then
                    MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                End If
       
                If Not IsNull(rsLIST!AccountNo) Then
                    MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                End If
                
                If Not IsNull(rsLIST!AsSigned) Then
                    MyList.SubItems(3) = (rsLIST!AsSigned)
                End If
                
                If Not IsNull(rsLIST!AgreementDate) Then
                    MyList.SubItems(4) = CStr(rsLIST!AgreementDate)
                End If
                
                If Not IsNull(rsLIST!SignedBy) Then
                    MyList.SubItems(5) = CStr(rsLIST!SignedBy)
                End If
                
                rsLIST.MoveNext
                
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub getContractToEdit()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "PlotNo", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 6.5
                .ListView1.ColumnHeaders.Add , , "Physical Location", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot P,ODASPAccount A, ODASMLeaseAgreement LA where  LA.PlotNo = P.PlotNo AND P.AccountNo = A.AccountNo and LA.ContractNo like '%" & CurrentRecord & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                DF = rsLIST.RecordCount
                Dim MyList As ListItem
                While Not rsLIST.EOF
                                                 
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!CommencementDate) Then
                                MyList.SubItems(2) = Format(rsLIST!CommencementDate, "dd/mm/yyyy")
                            End If
                            
                            If Not IsNull(rsLIST!expirydate) Then
                                MyList.SubItems(3) = Format(rsLIST!expirydate, "dd/mm/yyyy")
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(4) = CStr(rsLIST!PhysicalLocation)
                            End If
                            If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(5) = CStr(rsLIST!CompanyName)
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
Public Sub showALLUNLeasedPLOTS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "BillBoard No", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Location", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "No Of Faces", .ListView1.Width / 4
                

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPlotMast, ODASPPlot, ODASPAccount Where ODASPPLotMast.ExpiryDate >= '" & Format(frmODASSitesOnRoadReserve.txtStartDate.Text, "yyyy/mm/dd") & "' and ODASPPLotMast.ExpiryDate <= '" & Format(frmODASSitesOnRoadReserve.txtLastDate.Text, "yyyy/mm/dd") & "' and ODASPPlotMast.PlotNo= ODASPPlot.PlotNo and (LeasePrepared = 'N' or LeasePrepared is null) and OwenedByClient ='N' and ODASMInstallment.AccountNo=ODASPAccount.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MastNo))
                    
                        If Not IsNull(rsLIST!PlotName) Then
                            MyList.SubItems(1) = CStr(rsLIST!PlotName)
                        End If
                        
                        If Not IsNull(rsLIST!MastDetails) Then
                            MyList.SubItems(2) = CStr(rsLIST!MastDetails)
                        End If
                        
                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(3) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!NoofSites) Then
                            MyList.SubItems(4) = CStr(rsLIST!NoofSites)
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

Function FillList(Domain As String, lv As ListView, Optional FindString As String = "") As Boolean
 Screen.MousePointer = vbHourglass
      '==================================================================
      '  Purpose:   to fill a ListView control with data from a table or
      '             query
      '  Arguments: a Domain which is the name of the table or query, and
      '             a ListView control object
      '  Returns:   A Boolean value to indicate if the function was
      '             successful
      '==================================================================

      Dim rs As ADODB.Recordset
      Dim intTotCount As Integer
      Dim intCount1 As Integer, intCount2 As Integer
      Dim colNew As ColumnHeader, NewLine As ListItem

      On Error GoTo Err_Man

        ' Clear the ListView control.
        lv.ListItems.Clear
        lv.ColumnHeaders.Clear
    
        ' Set Variables.
         
        Set rs = New ADODB.Recordset
        cnCOMMON.CursorLocation = adUseClient
        rs.Open Domain, cnCOMMON, adOpenStatic, adLockOptimistic
       
        If Trim(FindString) = "" Then
        Else
                Dim strFilterString
                strFilterString = ""
            
                'Build filter string
                For i = 0 To rs.Fields.Count - 1
    
'                        If rs.Fields(i).Type = 202 Then
                                        strFilterString = strFilterString & "[" & rs.Fields(i).Name & "] like '%" & FindString & "%' " & " OR "
'                        End If
                        
                Next i
                'remove the last part of the string " OR "
                strFilterString = Left(strFilterString, Len(strFilterString) - Len(" OR "))
                
                rs.Filter = strFilterString
        End If
      
        ' Set Column Headers.
        For intCount1 = 0 To rs.Fields.Count - 1
             Set colNew = lv.ColumnHeaders.Add(, , rs(intCount1).Name, 1850)
        Next intCount1
        lv.View = lvwReport    ' Set View property to 'Report'.
    
        If rs.EOF Or rs.BOF Then
                    lv.View = lvwList
                    Set NewLine = lv.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                    lv.Enabled = True
                    Set rs = Nothing: Set NewLine = Nothing:  Screen.MousePointer = vbDefault: Exit Function
                    
        End If
        lv.Enabled = True
        ' Set Total Records Counter.
        rs.MoveLast
        intTotCount = rs.AbsolutePosition
        rs.MoveFirst
        If intTotCount = -1 Then
                rs.MoveFirst
                intTotCount = 0
                While Not rs.EOF
                        intTotCount = intTotCount + 1
                        DoEvents
                rs.MoveNext
                Wend
        End If
        
        rs.MoveFirst
          
        ' Loop through recordset and add Items to the control.
        For intCount1 = 1 To intTotCount
                    If IsNumeric(rs(0).Value) Then
                        Set NewLine = lv.ListItems.Add(, , CStr(rs(0).Value))
                    Else
                        Set NewLine = lv.ListItems.Add(, , rs(0).Value)
                    End If
                          
                    For intCount2 = 1 To rs.Fields.Count - 1
                            If Not IsNull(rs(intCount2)) Then
                                    NewLine.SubItems(intCount2) = rs(intCount2).Value
                            End If
                    Next intCount2
    
                    rs.MoveNext
                DoEvents
        Next intCount1
        
        cnCOMMON.CursorLocation = adUseServer
        
        If lv.ListItems.Count = 1 Then
            lv.ListItems(1).Checked = True
        End If
        
 Screen.MousePointer = vbDefault
    Exit Function

Err_Man:
         ' Ignore Error 94 which indicates you passed a NULL value.
         If err = 94 Then
            Resume Next
         Else
         ' Otherwise display the error message.
            MsgBox "Error: " & err.Number & Chr(13) & _
               Chr(10) & err.Description
         End If
Screen.MousePointer = vbDefault
      End Function
      
      
Public Sub checkOne(Item, lstView As ListView)
        Dim i, j As Double
        
        If Item.Checked = True Then
                    j = lstView.ListItems.Count
                    
                    If j = 0 Then Exit Sub
                    
                    For i = 1 To j
                                If lstView.ListItems(i) <> Item Then
                                   lstView.ListItems(i).Checked = False
                                End If
                    Next i
        Else
                    Item.Checked = False
        End If
End Sub

Public Sub checkAll(lstView As ListView)
        Dim i, j As Double
        
        j = lstView.ListItems.Count
                    
        If j = 0 Then Exit Sub
                    
        For i = 1 To j
                    lstView.ListItems(i).Checked = True
        Next i
End Sub

Public Sub UnCheckAll(lstView As ListView)
        Dim i, j As Double
        
        j = lstView.ListItems.Count
                    
        If j = 0 Then Exit Sub
                    
        For i = 1 To j
                    lstView.ListItems(i).Checked = False
        Next i
End Sub

Function SortListViewColumn(lv As Object, ColumnHeader)
'Check if the Sortkey is the same a the current one
    If lv.SortKey <> ColumnHeader.Index - 1 Then
        'When a column is clicked set the sortkey
        'to the columnheader index -1
        lv.SortKey = ColumnHeader.Index - 1
        lv.SortOrder = lvwAscending
    Else
        'If the column is already selected then change the
        'sortorder to be the opposite of what is currently
        'being used
        lv.SortOrder = IIf(lv.SortOrder = lvwAscending, _
                                lvwDescending, lvwAscending)
    End If
    
    'Set the sorted property to use the new sortkey
    'and sort the contents
    lv.Sorted = True
End Function

'Procedure used to search in listview
Public Sub search_in_listview(ByRef sListView As ListView, ByVal sFindText As String)
    Dim tmp_listtview As ListItem
    Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem)
    If Not tmp_listtview Is Nothing Then
        tmp_listtview.EnsureVisible
        tmp_listtview.Selected = True
    End If
End Sub


