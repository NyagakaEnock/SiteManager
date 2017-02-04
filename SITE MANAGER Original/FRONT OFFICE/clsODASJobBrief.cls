VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASJobBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rsTASK As ADODB.Recordset, strTASK As String

Public Sub generateTASKS()
On Error GoTo err
        
        Dim rsTASK As ADODB.Recordset, strTASK As String
        Set rsTASK = New Recordset
        
        strTASK = "SELECT * FROM ODASMJobBriefItems WHERE JobBriefNo = '" & CurrentRecord & "'"
        rsTASK.Open strTASK, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsTASK.EOF Or rsTASK.BOF Then Exit Sub
        
        Do While Not rsTASK.EOF
                
                Dim rsDEPT As ADODB.Recordset, strDEPT As String
                Set rsDEPT = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPMediaTask WHERE MediaCode = '" & rsTASK!MediaCode & "'"
                rsDEPT.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                If rsDEPT.EOF Or rsDEPT.BOF Then

                Else
                        Do While Not rsDEPT.EOF
                            
                                Set rsSAVE = New Recordset
                                strSQL = "SELECT * FROM ODASMJobCARD WHERE JobCardNo = '" & CurrentRecord & "' and DepartmentCode = '" & rsDEPT!DepartmentCode & "'"
                                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                                If rsSAVE.EOF Or rsSAVE.BOF Then
                                        rsSAVE.AddNew
                                        rsSAVE!JobCardNo = CurrentRecord
                                        rsSAVE!DepartmentCode = rsDEPT!DepartmentCode
                                        rsSAVE!dateprepared = Date
                                        rsSAVE!Preparedby = CurrentUserName
                                        rsSAVE!Status = "START"
                                        rsSAVE!Prepared = "Y"
                                        rsSAVE.Update
                                End If
                            
                            rsDEPT.MoveNext
                        Loop
                End If
                
                rsTASK.MoveNext
        Loop

Exit Sub

err:
    ErrorMessage
End Sub



Public Sub openJOBBRIEF()
On Error GoTo err
        
        Dim rsOPEN As ADODB.Recordset, strOPEN As String
        
        Set rsCHECK = New Recordset
        
        strCHECK = "SELECT * FROM ODASMJobCard WHERE JobCardNo = '" & CurrentRecord & "' and Opened ='N' "
        rsCHECK.Open strCHECK, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCHECK.EOF Or rsCHECK.BOF Then
                
                '/update the Job Brief File to Indicate that the
                '/ Job Cards are open
                
                Set rsOPEN = New Recordset
                
                strOPEN = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & CurrentRecord & "'"
                rsOPEN.Open strOPEN, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                If rsOPEN.EOF Or rsOPEN.BOF Then Exit Sub
                    rsOPEN!Status = "OPENED"
                    rsOPEN!opened = "Y"
                    rsOPEN!Closed = "N"
                    rsOPEN!DateOpened = Date
                    rsOPEN!openedBy = CurrentUserName
                    rsOPEN.Update
                    rsOPEN.Requery
        End If
        
strOPEN = Empty

Exit Sub

err:
    ErrorMessage
End Sub





Public Sub startJOBBRIEF()
On Error GoTo err
        
        Dim rsSTART As ADODB.Recordset, strSTART As String
        Set rsSTART = New ADODB.Recordset
        
        strSTART = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & CurrentRecord & "'"
        rsSTART.Open strSTART, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsSTART.EOF Or rsSTART.BOF Then Exit Sub
        
        If rsSTART.EOF Or rsSTART.BOF Then Exit Sub
        rsSTART!AuthorizedBy = CurrentUserName
        rsSTART!DateAuthorized = Date
        rsSTART!Authorized = "Y"
        rsSTART!Status = "STARTED"
        rsSTART!Started = "Y"
        rsSTART!DateStarted = Date
        rsSTART!StartedBy = CurrentUserName
        rsSTART.Update
        rsSTART.Requery
        

'rsSTART.Close

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub closeJOBBRIEF()
On Error GoTo err
        
        Set rsSAVE = New Recordset
        
        strSQL = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & CurrentRecord & "'"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub

        Set rsCONTROL = New Recordset
        
        strCONTROL = "SELECT * FROM ODASMJobCard WHERE JobCardNo = '" & CurrentRecord & "' and Closed = 'N'"
        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCONTROL.EOF Or rsCONTROL.BOF Then
                
                        rsSAVE!Status = "JOB BRIEF-CLOSED"
                        rsSAVE!Closed = "Y"
                        rsSAVE!DateClosed = Date
                        rsSAVE!ClosedBy = CurrentUserName
                        rsSAVE.Update
                        rsSAVE.Requery
        End If

rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub invoiceJOBBRIEF()
On Error GoTo err
        
        Set rsCONTROL = New Recordset
        
        strCONTROL = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & CurrentRecord & "'"
        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
        
        If rsSAVE.EOF Or rsSAVE.BOF Then
                rsSAVE!Status = "INVOICE"
                rsSAVE!Invoiced = "Y"
                rsSAVE!DateINVOICED = Date
                rsSAVE!InvoicedBy = CurrentUserName
                rsSAVE.Update
                rsSAVE.Requery
        End If

rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub GenerateJobItemNo()
On Error GoTo err

        Set rsSAVE = New ADODB.Recordset
        
        strSQL = "SELECT * FROM ODASMJobBriefItems WHERE JobBriefNo = '" & frmODASMJobBrief.txtJobBriefNo.Text & "'"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsSAVE
                If .BOF Or .EOF = True Then
                        frmODASMJobBrief.txtJobBriefItemNo.Text = 1
                 Else
                        Screen.ActiveForm.txtJobBriefItemNo.Text = .RecordCount + 1
                End If
                        
                Select Case Len(Trim(frmODASMJobBrief.txtJobBriefItemNo))
                        Case 1: frmODASMJobBrief.txtJobBriefItemNo.Text = Trim(frmODASMJobBrief.txtJobBriefNo) + "-0" + Trim(frmODASMJobBrief.txtJobBriefItemNo)
                        Case 2: frmODASMJobBrief.txtJobBriefItemNo.Text = Trim(frmODASMJobBrief.txtJobBriefNo) + "-" + Trim(frmODASMJobBrief.txtJobBriefItemNo)
                End Select
                
'                !JobBriefItemNo = frmODASMJobBrief.txtJobBriefNo.Text
'                .Update
'                .Requery

        End With
        
rsSAVE.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub saveJobBrief()
On Error GoTo err
       With frmODASMJobBrief
            
                strSQL = "select * from ODASMJobBrief Where JobBriefNo = '" & frmODASMJobBrief.txtJobBriefNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobBriefNo = Trim(.txtJobBriefNo.Text)
                        rsSAVE!CreatedBy = CurrentUserName
                        rsSAVE!DateCreated = Date
                        rsSAVE!Created = "Y"
                        rsSAVE!Approved = "N"
                        rsSAVE!Authorized = "N"
                        rsSAVE!Closed = "N"
                        rsSAVE!items = 0
                        rsSAVE!opened = "N"
                        rsSAVE!Invoiced = "N"
                        rsSAVE!Cancelled = "N"
                        rsSAVE!Started = "N"
                        rsSAVE!Status = "JOBBRIEF-CREATED"
                End If
                
                
                rsSAVE!descriptionOfOrder = .txtDescriptionOfOrder.Text
                rsSAVE!ProductCode = .txtProductCode.Text
                rsSAVE!ReceivedBy = .txtReceivedBy.Text
                rsSAVE!BriefBy = .txtBriefBy.Text
                rsSAVE!JobBriefDate = .txtJobBriefDate.Text
                rsSAVE!expectedDOC = .txtExpectedDOC.Text
                rsSAVE!QuotationNumber = .txtQuotationNo.Text
                rsSAVE!Accountno = .txtAccountNo.Text
                rsSAVE!deadlineDate = .txtDeadlineDate.Text
                rsSAVE!AccPeriod = CurrentPeriod
                rsSAVE!AdditionalInformation = .txtAdditionalInformation.Text
                
                If .chkAcquireSite.Value = 1 Then
                        rsSAVE!AcquireSite = "Y"
                Else: rsSAVE!AcquireSite = "N"
                End If
                
                rsSAVE.Update
                rsSAVE.Requery
                
                Set rsNewRecord = Nothing
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub saveJobBriefItems()
On Error GoTo err
    
       With frmODASMJobBrief
                Dim strIlluminated, strBordered As String
            
                If .chkBorder.Value = True Then strBordered = "Y" Else strBordered = "N"
                If .chkIlluminate.Value = True Then strIllumintaed = "Y" Else strIlluminated = "N"

                strSQL = "select * from ODASMJobBriefItems Where JobBriefNo = '" & frmODASMJobBrief.txtJobBriefNo & "' and JobBriefItemNo = '" & frmODASMJobBrief.txtJobBriefItemNo.Text & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobBriefNo = .txtJobBriefNo.Text
                        rsSAVE!JobBriefItemNo = .txtJobBriefItemNo.Text
                        rsSAVE!CreatedBy = CurrentUserName
                        rsSAVE!DateCreated = Date
                End If
                        rsSAVE!MediaCode = .txtMediaCode.Text
                        rsSAVE!MediaSize = .txtMediaSize.Text
                        rsSAVE!SiteNo = .txtSiteNo.Text
                        rsSAVE!illuminated = strIlluminated
                        rsSAVE!SidingCode = .cboSidingCode.Text
                        rsSAVE!ColorCode = .cboColorCode.Text
                        rsSAVE!itemQuantity = Val(.txtItemQuantity.Text)
                        rsSAVE!PhysicalLocation = .txtPhysicalLocation.Text
                        rsSAVE!Bordered = strBordered
                        rsSAVE.Update
                        rsSAVE.Requery
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub saveJobBriefATTACHMENT()
On Error GoTo err
    
       With frmODASMJobBrief

                strSQL = "select * from ODASMJobBriefAttachment Where JobBriefNo = '" & frmODASMJobBrief.txtJobBriefNo & "' and Attachmentcode = '" & .txtAttachmentCode.Text & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobBriefNo = .txtJobBriefNo.Text
                        rsSAVE!AttachmentCode = .txtAttachmentCode.Text
                        rsSAVE!Preparedby = CurrentUserName
                        rsSAVE!dateprepared = Date
                End If
                
                rsSAVE.Update
                rsSAVE.Requery
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub updateSITE()
On Error GoTo err
    
        ' procedure is used to update the site on the Bill Board. If the Bill Board has two sides
        ' then both are updared to show that they are available
        
        
      With frmODASMJobBrief
                
                strCONTROL = "select * from ODASPPlotSite, ODASPPlot Where ODASPPlotSite.SiteNo = '" & .txtSiteNo & "' and ODASPPlotSite.status = 'SITE-AVAILABLE' and ODASPPlotSite.PlotNo = ODASPPlot.PlotNo;"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Do While Not rsCONTROL.EOF
                
                        Set rsSAVE = New ADODB.Recordset
                        strSQL = "select * from ODASPPlotSite Where ODASPPlotSite.PlotNo = '" & rsCONTROL!PlotNo & "' and ODASPPlotSite.status = 'SITE-AVAILABLE';"
                        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        If rsSAVE.EOF Or rsSAVE.BOF Then
                        Else
                                rsSAVE!Status = "SITE-RESERVED"
                                rsSAVE!JobBriefItemNo = .txtJobBriefItemNo
                                rsSAVE!JobBriefNo = .txtJobBriefItemNo.Text
                                rsSAVE!ReservationDate = .txtJobBriefDate.Text
                                rsSAVE.Update
                        End If
                        rsCONTROL.MoveNext
                Loop
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub



Public Sub updateMAST()
On Error GoTo err
    
        ' procedure is used to update the mast which the Bill Board is installed on. If the Bill Board has two sides
        ' then both are updared to show that they are available
        
        
    With frmODASMJobBrief
                
            Set rsSAVE = New ADODB.Recordset
            strSQL = "select * from ODASPPlotMast Where ODASPPlotMast.MastNo = '" & .txtSiteNo.Text & "' and ODASPPlotMAST.status = 'SITE-AVAILABLE';"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
            If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
            
                rsSAVE!Status = "SITE-RESERVED"
                rsSAVE!JobBriefItemNo = .txtJobBriefItemNo
                rsave!JobBriefNo = .txtJobBriefItemNo.Text
                rsSAVE!ReservationDate = .txtJobBriefDate.Text
                rsSAVE.Update
                rsSAVE.Requery
                
            Set rsSAVE = Nothing
            strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableRECORD()
On Error GoTo err
        With frmODASMJobBrief
            .txtAdditionalInformation.Locked = False
            .txtCompanyName.Locked = True
            .txtBriefBy.Locked = True
            .txtProductCode.Locked = True
            .txtReceivedBy.Locked = True
            .txtDescriptionOfOrder.Locked = False
            .txtExpectedDOC.Locked = False
            .txtItemQuantity.Locked = False
            .txtMediaCode.Locked = False
            .txtMediaSize.Locked = False
            .cboColorCode.Locked = False
            .cboSidingCode.Locked = False
            .chkAcquireSite.Enabled = True
            .chkBorder.Enabled = True
            .chkIlluminate.Enabled = True
            .txtBriefBy.Locked = False
            .txtProductCode.Locked = False
            .txtReceivedBy.Locked = False
            .txtAttachmentCode.Locked = False
            .dtpDeadlineDate.Enabled = True
            .DTPickerExpectedDOC.Enabled = True
            .txtSiteLocation.Locked = False
            .txtPhysicalLocation.Locked = False
            
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Function ValidRecord() As Boolean
With frmODASMJobBrief
    
    If .txtContactName.Text = Empty Then
        strMessage = "Name Of Contact Person Required...!"
        .txtContactName.SetFocus
    
    ElseIf .cboSidingCode.Text = Empty Then
        strMessage = "Advertisement siding required...!"
        .cboSidingCode.SetFocus
    
    
    ElseIf .txtBriefBy.Text = Empty Then
        strMessage = "Briefed by cannot be Left Blank"
        .txtBriefBy.SetFocus
    
    ElseIf .txtProductCode.Text = Empty Then
        strMessage = "The Product Code Cannot be Blank"
        .txtProductCode.SetFocus
    
    ElseIf .txtReceivedBy.Text = Empty Then
        strMessage = "Received By Cannot be Blank"
        .txtReceivedBy.SetFocus
        
    ElseIf .txtJobBriefDate.Text = Empty Then
        Strmsssage = "The Job Brief Date is Blank"
        .txtJobBriefDate.SetFocus
    
    ElseIf .txtAccountNo.Text = Empty Then
        strMessage = "The Customer Has to be Selected"
        .txtAccountNo.SetFocus
    
    ElseIf .txtDescriptionOfOrder.Text = Empty Then
        strMessage = "The Description of Order is Required to Identify the Order"
        .txtDescriptionOfOrder.SetFocus
    
    ElseIf .txtDeadlineDate.Text = Empty Then
        strMessage = "The Deadline Date is Required"
        .txtDeadlineDate.SetFocus
    
    ElseIf .txtAttachmentCode.Text = Empty Then
        strMessage = "The Attachment Code is Necessary"
        .txtAttachmentCode.SetFocus

    ElseIf .txtExpectedDOC.Text = Empty Then
        strMessage = "The Expected Commencement Date is Necessary"
        .txtExpectedDOC.SetFocus
    
    ElseIf DateDiff("D", .txtExpectedDOC, .txtDeadlineDate) < 0 Then
        strMessage = "The expected DOC cannot be after the deadline Date"
        .txtDeadlineDate.SetFocus
    
    ElseIf .txtMediaCode.Text = Empty Then
        strMessage = "The Media Code cannot be Left Blank"
        .txtMediaCode.SetFocus
    
    ElseIf .txtMediaSize.Text = Empty Then
        strMessage = "The Media Size cannot be left Blank "
        .txtMediaSize.SetFocus
    
    
    ElseIf CDbl(.txtItemQuantity.Text) <= 0 Then
        strMessage = "The Item Quantity is Required"
        .txtItemQuantity.SetFocus
    
    Else
        ValidRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
End Function


Public Sub clearRECORD()
On Error GoTo err
        
        With frmODASMJobBrief
                .txtMediaSize.Text = ""
                .chkIlluminate.Value = 0
                .cboSidingCode.Text = Empty
                .chkBorder.Value = 0
                .txtItemQuantity.Text = 1
                .txtSiteName.Text = Empty
                .txtPhysicalLocation.Text = Empty
                .txtAttachmentCode.Text = Empty
                .txtJobBriefItemNo.Text = Empty
                .txtMediaSize.Text = Empty
                .txtSiteNo.Text = Empty
                .ListView3.ListItems.Clear
'                RemoveCurrentList3Item
'                RemoveCurrentList2Item
            
        End With

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub loadCONTACT()
On Error GoTo err
        
        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT * FROM ODASPAccount WHERE AccountNo = '" & Screen.ActiveForm.txtAccountNo.Text & "' ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
                    If .BOF And .EOF Then Exit Sub
                    Screen.ActiveForm.txtCompanyName.Text = !CompanyName
                    Screen.ActiveForm.txtContactName.Text = !ContactPerson
        End With

rsCONTROL.Close
strSQL = ""

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub loadDEFAULTS()
On Error GoTo err
    
    With frmODASMJobBrief
        .txtJobBriefDate.Text = Date
        .txtDeadlineDate.Text = Date
        .txtExpectedDOC.Text = Date
        .dtpBriefDate.Value = Date
        .dtpDeadlineDate.Value = Date
        .DTPickerExpectedDOC.Value = Date
        .txtItemQuantity.Text = 1
        .txtReceivedBy.Text = CurrentUserName
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadCATEGORY()
On Error GoTo err
        
        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT * FROM AdvertCategories WHERE MediaCode = '" & Screen.ActiveForm.txtMediaCode.Text & "' ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
                    If .BOF And .EOF Then Exit Sub
                    Screen.ActiveForm.txtCategoryName.Text = !CategoryName
        End With

rsCONTROL.Close
strSQL = ""

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub calculatePRICE()
On Error GoTo err
        
        With frmODASMJobBrief
                
                .txtItemQuantity.Text = 1
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calculateTOTALQUANTITY()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT sum(ItemQuantity) as GrandQuantity FROM ODASMJobBriefItems WHERE JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then
                frmODASMJobBrief.txtTotalQuantity.Text = 0
        ElseIf IsNull(rsCONTROL!GrandQuantity) = True Then
                frmODASMJobBrief.txtTotalQuantity.Text = 0
        Else: frmODASMJobBrief.txtTotalQuantity.Text = rsCONTROL!GrandQuantity
        End If


Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calculateTOTALPRICE()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT sum(NetItemPrice) as GrandPrice FROM ODASMJobBriefItems WHERE JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then
                frmODASMJobBrief.txtTotalPrice.Text = 0
        ElseIf IsNull(rsCONTROL!GrandPrice) = True Then
                frmODASMJobBrief.txtTotalPrice.Text = 0
        Else: frmODASMJobBrief.txtTotalPrice.Text = rsCONTROL!GrandPrice
        End If


Exit Sub

err:
    ErrorMessage
End Sub


Public Sub loadSiteName()
On Error GoTo err
    With frmODASMJobBrief
        
        Set rsCONTROL = New ADODB.Recordset
        
        If rsMEDIA!RequireBillBoard = "Y" Then
                strSQL = "SELECT * FROM ODASPPlotSite WHERE SiteNo = '" & .txtSiteNo.Text & "' ;"
        ElseIf rsMEDIA!RequireSite = "Y" Then
             strSQL = "SELECT * FROM ODASPPlot, ODASPPlotSite WHERE ODASPPlotSite.SiteNo = '" & .txtSiteNo.Text & "' and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo ;"
        End If
        
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCONTROL.BOF And rsCONTROL.EOF Then Exit Sub
        
            If rsMEDIA!RequireBillBoard = "Y" Then
                    .txtSiteName.Text = rsCONTROL!SiteDetails
            ElseIf rsMEDIA!RequireSite = "Y" Then
                    .txtSiteName.Text = rsCONTROL!PlotName
            End If
            
    End With

rsCONTROL.Close
strSQL = ""

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub loadMEDIADETAILS()
On Error GoTo err
        
        Set rsMEDIA = New ADODB.Recordset
        
        strSQL = "SELECT * FROM ODASPMedia WHERE MediaCode = '" & Screen.ActiveForm.txtMediaCode.Text & "' ;"
        rsMEDIA.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsMEDIA
                    If .BOF And .EOF Then Exit Sub
                    
                    If !LocDepOnSize = "Y" Then
                            'showALLMEDIASIZES3
                    Else:
                            frmODASMJobBrief.ListView3.ListItems.Clear
                            frmODASMJobBrief.txtMediaSize.Text = "N/A"
                    End If
                    
                    If !RequireSite = "Y" Then
                            'showALLAvailableSites
                    Else:
                            frmODASMJobBrief.txtSiteNo.Text = "N/A"
                            frmODASMJobBrief.txtSiteName.Text = "NOT APPLICABLE"
                    End If
                    
                    If !RequireBillBoard = "Y" Then
                            frmODASMJobBrief.FrameSites.Caption = "Bill Board Available"
                            'showALLMASTS
                    End If
                    
                    If !RequireNothing = "Y" Then
                            frmODASMJobBrief.FrameSites.Caption = "No Site Required"
                            frmODASMJobBrief.ListView4.ListItems.Clear
                            frmODASMJobBrief.txtSiteNo.Text = "N/A"
                            frmODASMJobBrief.txtSiteName.Text = "NOT APPLICABLE"
                    End If
                    
        End With

strSQL = ""

Exit Sub
err:
    ErrorMessage
End Sub

