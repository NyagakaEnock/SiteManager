VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdminCentre 
   Caption         =   "Resource Scheduling -ADMININSTRATION CENTRE"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmAdminCentre.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   4455
      Left            =   5640
      TabIndex        =   16
      Top             =   720
      Width           =   6135
      Begin MSComctlLib.ListView ListView2 
         Height          =   4215
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7435
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Width           =   11895
      Begin MSComctlLib.ListView ListView1 
         Height          =   2775
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   4895
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Functions"
      Height          =   4455
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   5535
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5160
         TabIndex        =   21
         Top             =   3240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   64356353
         CurrentDate     =   38274
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Top             =   2640
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   64356353
         CurrentDate     =   38274
      End
      Begin VB.TextBox txtNewEndDate 
         BackColor       =   &H00FFC0C0&
         Height          =   405
         Left            =   3000
         TabIndex        =   19
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox txtNewStartDate 
         BackColor       =   &H00FFC0C0&
         Height          =   405
         Left            =   3000
         TabIndex        =   18
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select A&ll"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdApprove 
         BackColor       =   &H008080FF&
         Caption         =   "&APPROVE"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3720
         Width           =   4935
      End
      Begin VB.OptionButton optChoices 
         Caption         =   "Reinstate Advertisement(s)"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   13
         Top             =   1680
         Width           =   2415
      End
      Begin VB.OptionButton optChoices 
         Caption         =   "Reinstate Site(s)"
         Height          =   255
         Index           =   7
         Left            =   3120
         TabIndex        =   12
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton optChoices 
         Caption         =   "Discontinue Site(s)"
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   11
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton optChoices 
         Caption         =   "Discontiue Advertisemt(s)"
         Height          =   375
         Index           =   4
         Left            =   480
         TabIndex        =   10
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optChoices 
         Caption         =   "Site Renewals"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optChoices 
         Caption         =   "Advertisement(s) Renewals"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   8
         Top             =   2280
         Width           =   2415
      End
      Begin VB.OptionButton optChoices 
         Caption         =   "Site(s) Approval"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optChoices 
         Caption         =   "Contract(s) Approvals"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Text            =   "***R.S (Plus) ADMINISTRATION CENTRE ***"
         Top             =   0
         Width           =   11895
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Sheet View"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5160
      Width           =   3135
   End
End
Attribute VB_Name = "frmAdminCentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PurchaseOrderNo, SerialNo, SiteNo, LordNo, Serial, Reinstatesiteno, DiscontinueSitenO, ReinstateAdvertNo, AdvertCode, SiteNum, PurchaseNo

Private Sub cmdApprove_Click()
'On Error GoTo Err
Dim i, j, k
With Me
If optChoices(0).Value = True Then
 j = .ListView1.ListItems.Count
   If j = 0 Or .ListView1.View <> lvwReport Then Exit Sub
   
   For i = 1 To j
     If .ListView1.ListItems(i).Checked = True Then
       SerialNo = .ListView1.ListItems(i).Text
       PurchaseOrderNo = .ListView1.ListItems(i).SubItems(1)
       
       Set rsLineUpdate = New ADODB.Recordset
        rsLineUpdate.Open "UPDATE AdvertContractRequisition SET ApprovedStatus = '" & "Y" & "' ,ApprovedBy = '" & CurrentUserName & "',DateApproved = '" & MyCurrentDate & "' WHERE PurchaseOrderNo = '" & PurchaseOrderNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
       Set rsLineUpdate = Nothing
       
       Set rsLineUpdate = New ADODB.Recordset
       rsLineUpdate.Open "UPDATE AdvertContractRequisitionData SET ApprovedStatus = '" & "Y" & "' ,ApprovedBy = '" & CurrentUserName & "',DateApproved = '" & MyCurrentDate & "' WHERE SerialNo = '" & SerialNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
       Set rsLineUpdate = Nothing
           
     End If
    Next i
    
      ElseIf .optChoices(1).Value = True Then
         j = .ListView1.ListItems.Count
           If j = 0 Or .ListView1.View <> lvwReport Then Exit Sub
           
           For i = 1 To j
             If .ListView1.ListItems(i).Checked = True Then
             SiteNo = .ListView1.ListItems(i).Text
             LordNo = .ListView1.ListItems(i).SubItems(17)
             
             Set rsEditRecord = New ADODB.Recordset
             rsEditRecord.Open "UPDATE AdvertSites SET Discontinued = '" & "N" & "',DateApproved = '" & MyCurrentDate & "',ApprovedBy = '" & CurrentUserName & "',ApprovedStatus = '" & "Y" & "',RenewalApprovalStatus = '" & "Y" & "' WHERE sITENO = '" & SiteNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
             Set rsEditRecord = Nothing
             
             Set rsEditRecord = New ADODB.Recordset
              rsEditRecord.Open "UPDATE AdvertSiteLords SET ApprovedStatus = '" & "Y" & "',ApprovedBy = '" & CurrentUserName & "',DateApproved = '" & MyCurrentDate & "' WHERE LandLordNo = '" & LordNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
             Set rsEditRecord = Nothing
             End If
           Next i
           
      ElseIf .optChoices(2).Value = True Then
           j = .ListView1.ListItems.Count
            If j = 0 Or .ListView1.View <> lvwReport Then Exit Sub
            
            For i = 1 To j
           If .ListView1.ListItems(i).Checked = True Then
            Serial = .ListView1.ListItems(i).Text
            PurchaseNo = .ListView1.ListItems(i).SubItems(1)
            
            Set rsEditRecord = New ADODB.Recordset
            rsEditRecord.Open "UDDATE AdvertContractRequisition SET RenewalApprovalStatus = '" & "Y" & "',RenewalApprovedBy = '" & CurrentUserName & "',DateOfRenewalApproval = '" & MyCurrentDate & "' WHERE PurchaseOrderNo = '" & PurchaseNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsEditRecord = Nothing
       
            Set rsEditRecord = New ADODB.Recordset
            rsEditRecord.Open "UDDATE AdvertContractRequisitionData SET RenewalApprovalStatus = '" & "Y" & "',RenewalApprovedBy = '" & CurrentUserName & "',DateOfRenewalApproval = '" & MyCurrentDate & "' WHERE SerialNo = '" & Serial & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsEditRecord = Nothing
           End If
         Next i
         
       ElseIf .optChoices(3).Value = True Then
          j = .ListView1.ListItems.Count
          If j = 0 Or .ListView1.View <> lvwReport Then Exit Sub
          
          For i = 1 To j
            If .ListView1.ListItems(i).Checked = True Then
            SiteNum = .ListView1.ListItems(i).Text
            
            If .txtNewStartDate.Text = "" And .txtNewEndDate.Text = "" Then
            MsgBox "New contract start date and end date required", vbExclamation, "Contract Renewal"
            .txtNewStartDate.SetFocus
            Else
            
            Set rsLineUpdate = New ADODB.Recordset
            rsLineUpdate.Open "UPDATE AdvertSites SET  RenewalApprovalStatus = '" & "Y" & "',RenewalApprovedBy = '" & CurrentUserName & "',RenewalDateApproved = '" & MyCurrentDate & "'WHERE SiteNo = '" & SiteNum & "',ConractFinish = '" & Format(.txtNewEndDate.Text, "MMMM dd,yy") & "',ContractStart = '" & Format(.txtNewStartDate.Text, "MMMM dd,yy") & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsLineUpdate = Nothing
            
            End If
          End If
            Next i
       
       ElseIf .optChoices(4).Value = True Then
          j = .ListView1.ListItems.Count
          If j = 0 Or .ListView1.View <> lvwReport Then Exit Sub
          
          For i = 1 To j
          If .ListView1.ListItems(i).Checked = True Then
           AdvertCode = .ListView1.ListItems(i).Text
           
           Set rsLineUpdate = New ADODB.Recordset
           rsLineUpdate.Open "UPDATE AdvertBBDetails SET Discontinued = NULL WHERE BillBoardNo = '" & AdvertCode & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
           Set rsLineUpdate = Nothing
          End If
          Next i
            
       ElseIf .optChoices(5).Value = True Then
          j = .ListView1.ListItems.Count
          If j = 0 Or .ListView1.View <> lvwReport Then Exit Sub
          
          For i = 1 To j
          If .ListView1.ListItems(i).Checked = True Then
            DiscontinueSitenO = .ListView1.ListItems(i).Text
            
            Set rsLineUpdate = New ADODB.Recordset
            rsLineUpdate.Open "UPDATE AdvertSites SET Discontinued = Null ,DiscontinuedBy = '" & CurrentUserName & "',DateDiscontinued = '" & MyCurrentDate & "' WHERE SitenO = '" & DiscontinueSitenO & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsLineUpdate = Nothing
          End If
         Next i
         
       ElseIf .optChoices(7).Value = True Then
         j = .ListView1.ListItems.Count
         If j = 0 Or .ListView1.View <> lvwReport Then Exit Sub
         
         For i = 1 To j
          If .ListView1.ListItems(i).Checked = True Then
           Reinstatesiteno = .ListView1.ListItems(i).Text
           
           Set rsLineUpdate = New ADODB.Recordset
           rsLineUpdate.Open "UPDATE AdvertSites SET Discontinued = '" & "N" & "',ReinstatedBy = '" & CurrentUserName & "',DateReinstated = '" & MyCurrentDate & "' WHERE sITENO = '" & Reinstatesiteno & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
           Set rsLineUpdate = Nothing
          End If
          Next i
          
      ElseIf .optChoices(8).Value = True Then
        j = .ListView1.ListItems.Count
        If j = 0 Or .ListView1.View <> lvwReport Then Exit Sub
        
        For i = 1 To j
        If .ListView1.ListItems(i).Checked = True Then
         ReinstateAdvertNo = .ListView1.ListItems(i).Text
         
         Set rsLineUpdate = New ADODB.Recordset
         rsLineUpdate.Open "UPDATE AdvertBBDetails SET Discontinued = '" & "N" & "',ReinstatedBy = '" & CurrentUserName & "',DateReinstated = '" & MyCurrentDate & "' WHERE Billboardno = '" & ReinstateAdvertNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
         Set rsLineUpdate = Nothing
        End If
        Next i
                      
End If
MsgBox "Requested Transaction Successfully Completed", vbInformation, "Resource Scheduling"
End With
Exit Sub
Err:
    ErrorMessage
End Sub



Private Sub DTPicker1_CloseUp()
If Not NewRecord And Not EditRecord Then Exit Sub
    If Me.DTPicker1.Value < Date Then
        MsgBox "Wrong Date! Contract Start Date Cannot be Earlier than the Date Today!!", vbCritical + vbOKOnly, "Invalid Date"
        Me.txtNewStartDate.Text = Empty: Me.txtNewStartDate.SetFocus
    Else
        Me.txtNewStartDate.Text = Me.DTPicker1.Value
        Me.txtNewEndDate.SetFocus
    End If
End Sub



Private Sub DTPicker2_CloseUp()
If Not NewRecord And Not EditRecord Then Exit Sub
    If Me.DTPicker2.Value < Date Then
        MsgBox "Wrong Date! Contract End Date Cannot be Earlier than the Date Today!!", vbCritical + vbOKOnly, "Invalid Date"
        Me.txtNewEndDate.Text = Empty: Me.txtNewEndDate.SetFocus
    Else
        Me.txtNewStartDate.Text = Me.DTPicker1.Value
        Me.cmdApprove.SetFocus
    End If
End Sub

Private Sub Form_Load()
With Me
Call ShowAdvertsStructure
Call GetUnApprovedContractsStructure
'.DTPicker1.
End With
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo Err

If Me.ListView2.ListItems.Count = 0 Or Me.ListView2.View <> lvwReport Then Item.Checked = False: Exit Sub
    
    Dim i, j, k
    j = Me.ListView2.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView2.ListItems(i).Text <> Item Then
            Me.ListView2.ListItems(i).Checked = False
        End If
    Next i
    
    If Me.ListView2.ColumnHeaders(1).Text = "Contract No" Then
        CurrentOrder = Item
       Call ShowAdvertsUnderCurrentContract
       If MsgBox("Do you want to print the contractform ?", vbYesNo, "Contract Approval") = vbYes Then
       SelectedProduct = CurrentOrder
       Load frmRPTContractRequisitionForm
       frmRPTContractRequisitionForm.Show 1, Me
       Else
        Me.ListView1.SetFocus
       End If
    ElseIf Me.ListView2.ColumnHeaders(1).Text = "Contract Number" Then
        CurrrentOrder1 = Item
        Call ShowRenewalAdvertsUnderCurrentContract
        Me.ListView1.SetFocus
        
        
    ElseIf Item.Checked = False Then
        Me.ListView2.ListItems.Clear
    End If
    
Exit Sub
Err:
    ErrorMessage
End Sub
Public Sub ShowAdvertsUnderCurrentContract()
'On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Contract No ", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Code ", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Type", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Length", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Width", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Duration", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Days", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Cost", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract End Date", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Approved", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Paid", .ListView1.Width / 6.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisitionData WHERE PurchaseOrderNo='" & Trim(CurrentOrder) & "' AND  ApprovedStatus IS NULL AND ContractEndDate > '" & Today & "'  ORDER BY SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(1) = CStr(rsLIST!PurchaseOrderNo)
    End If
    
    If Not IsNull(rsLIST!AdvCode) Then
        MyList.SubItems(2) = CStr(rsLIST!AdvCode)
    End If
    
    If Not IsNull(rsLIST!AdvName) Then
        MyList.SubItems(3) = CStr(rsLIST!AdvName)
    End If
    
    If Not IsNull(rsLIST!AdvType) Then
        MyList.SubItems(4) = CStr(rsLIST!AdvType)
    End If
    
    If Not IsNull(rsLIST!AdvLength) Then
        MyList.SubItems(5) = CStr(rsLIST!AdvLength)
    End If
    
    If Not IsNull(rsLIST!AdvWidth) Then
        MyList.SubItems(6) = CStr(rsLIST!AdvWidth)
    End If
    
    If Not IsNull(rsLIST!Duration) Then
        MyList.SubItems(7) = CStr(rsLIST!Duration)
    End If
    
    If Not IsNull(rsLIST!Days) Then
        MyList.SubItems(8) = CStr(rsLIST!Days)
    End If
    
            
    If Not IsNull(rsLIST!AdvCost) Then
        MyList.SubItems(9) = FormatNumber(rsLIST!AdvCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!ContractStartDate) Then
        MyList.SubItems(10) = CStr(rsLIST!ContractStartDate)
    End If
    
    If Not IsNull(rsLIST!ContractEndDate) Then
        MyList.SubItems(11) = CStr(rsLIST!ContractEndDate)
    End If
    
        
    If IsNull(rsLIST!ApprovedStatus) Then
        MyList.SubItems(12) = CStr("NO")
    ElseIf Not IsNull(rsLIST!ApprovedStatus) Then
        If rsLIST!ApprovedStatus = "Y" Then
            MyList.SubItems(12) = CStr("YES")
        Else
            MyList.SubItems(12) = CStr("NO")
        End If
    End If
    
    
    If IsNull(rsLIST!PaidStatus) Then
        MyList.SubItems(13) = CStr("NO")
    ElseIf Not IsNull(rsLIST!PaidStatus) Then
        If rsLIST!PaidStatus = "Y" Then
            MyList.SubItems(13) = CStr("YES")
        Else
            MyList.SubItems(13) = CStr("NO")
        End If
    End If
    
    
    rsLIST.MoveNext
    
Wend

'.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
'.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub ShowRenewalAdvertsUnderCurrentContract()
'On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Contract No ", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Code ", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Type", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Length", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Width", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Duration", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Days", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Cost", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract End Date", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Approved", .ListView1.Width / 6.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisitionData WHERE PurchaseOrderNo='" & Trim(CurrentOrder) & "' AND  RenewalApprovalStatus IS NULL ORDER BY SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(1) = CStr(rsLIST!PurchaseOrderNo)
    End If
    
    If Not IsNull(rsLIST!AdvCode) Then
        MyList.SubItems(2) = CStr(rsLIST!AdvCode)
    End If
    
    If Not IsNull(rsLIST!AdvName) Then
        MyList.SubItems(3) = CStr(rsLIST!AdvName)
    End If
    
    If Not IsNull(rsLIST!AdvType) Then
        MyList.SubItems(4) = CStr(rsLIST!AdvType)
    End If
    
    If Not IsNull(rsLIST!AdvLength) Then
        MyList.SubItems(5) = CStr(rsLIST!AdvLength)
    End If
    
    If Not IsNull(rsLIST!AdvWidth) Then
        MyList.SubItems(6) = CStr(rsLIST!AdvWidth)
    End If
    
    If Not IsNull(rsLIST!Duration) Then
        MyList.SubItems(7) = CStr(rsLIST!Duration)
    End If
    
    If Not IsNull(rsLIST!Days) Then
        MyList.SubItems(8) = CStr(rsLIST!Days)
    End If
    
            
    If Not IsNull(rsLIST!AdvCost) Then
        MyList.SubItems(9) = FormatNumber(rsLIST!AdvCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!ContractStartDate) Then
        MyList.SubItems(10) = CStr(rsLIST!ContractStartDate)
    End If
    
    If Not IsNull(rsLIST!ContractEndDate) Then
        MyList.SubItems(11) = CStr(rsLIST!ContractEndDate)
    End If
    
        
    If IsNull(rsLIST!ApprovedStatus) Then
        MyList.SubItems(12) = CStr("NO")
    ElseIf Not IsNull(rsLIST!ApprovedStatus) Then
        If rsLIST!ApprovedStatus = "Y" Then
            MyList.SubItems(12) = CStr("YES")
        Else
            MyList.SubItems(12) = CStr("NO")
        End If
    End If
           
    rsLIST.MoveNext
    
Wend

'.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
'.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub ShowAdvertsStructure()
'On Error GoTo Err
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Advert Code ", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Type", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Length", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Width", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Duration", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Days", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Cost", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract End Date", .ListView1.Width / 5.4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Approved", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Paid", .ListView1.Width / 6.5 ', lvwColumnCenter

.ListView1.View = lvwReport
End With
Exit Sub
Error:
   ErrorMessage
End Sub

Private Sub optChoices_Click(Index As Integer)
'On Error GoTo Err
With Me
'If Not NewRecord And Not EditRecord Then Exit Sub
Select Case Index
Case 0
    Call GetUnApprovedContracts
    Me.optChoices(0).Value = True
   
Case 1
    Me.ListView2.ListItems.Clear
    Call GetUnApprovedSites
    Me.optChoices(1).Value = True
    
Case 2
     Call GetUnApprovedRenewalContracts
     Me.optChoices(2).Value = True
Case 3
     Call GetRenewalApprovalSites
     Me.optChoices(3).Value = True
Case 4
     Call GetNonDiscontinuedAdverts
     Me.optChoices(4).Value = True
Case 5
     Call GetNonDiscontinuedSites
     Me.optChoices(5).Value = True

Case 7
    Call DiscontinuedSites
    Me.optChoices(7).Value = True

Case 8
    Call GetDiscontinuedAdverts
    Me.optChoices(8).Value = True
Case Else
    Exit Sub
End Select
Exit Sub
Err:
    ErrorMessage
End With
End Sub
Private Sub GetDiscontinuedAdverts()

'On Error GoTo Err
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Advert Code ", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Type", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Length", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Width", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Weight", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Duration", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Days", .ListView1.Width / 5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertBBDetails  WHERE Discontinued IS NULL  AND BillBoardNo IS NOT NULL  ORDER BY BillBoardNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!BillBoardNo))

    If Not IsNull(rsLIST!Name) Then
        MyList.SubItems(1) = CStr(rsLIST!Name)
    End If
    
    If Not IsNull(rsLIST!TypeName) Then
        MyList.SubItems(2) = Trim(CStr(rsLIST!TypeName))
    End If
    
    If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!Length))
    End If
    
    If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(4) = Trim(CStr(rsLIST!Width))
    End If
    
    If Not IsNull(rsLIST!Weight) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!Weight))
    End If
    
    If Not IsNull(rsLIST!DurationName) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!DurationName))
    End If
    
    If Not IsNull(rsLIST!NoOfDays) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!NoOfDays))
    End If
    
    
            
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub



Private Sub DiscontinuedSites()
'On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Site Name ", .ListView1.Width / 4.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Billboard Number", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 1 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Allocated", .ListView1.Width / 1 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Valid", .ListView1.Width / 1 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertSites WHERE  Discontinued  IS NULL ", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))

    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(1) = CStr(rsLIST!SiteName)
    End If
    
    If Not IsNull(rsLIST!BBNo) Then
        MyList.SubItems(2) = CStr(rsLIST!BBNo)
    End If
    
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(3) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!sitephysicalAddress) Then
        MyList.SubItems(4) = CStr(rsLIST!sitephysicalAddress)
    End If
    
               
    If IsNull(rsLIST!AllocationStatus) Then
        MyList.SubItems(5) = CStr("NO")
    ElseIf Not IsNull(rsLIST!AllocationStatus) Then
        If rsLIST!AllocationStatus = "Y" Then
            MyList.SubItems(5) = CStr("YES")
        Else
            MyList.SubItems(5) = CStr("NO")
        End If
    End If
    
    
    If IsNull(rsLIST!Validstatus) Then
        MyList.SubItems(6) = CStr("NO")
    ElseIf Not IsNull(rsLIST!Validstatus) Then
        If rsLIST!Validstatus = "Y" Then
            MyList.SubItems(6) = CStr("YES")
        Else
            MyList.SubItems(6) = CStr("NO")
        End If
    End If
    
    
    rsLIST.MoveNext
    
Wend

'.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
'.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub GetNonDiscontinuedSites()
'On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Site Name ", .ListView1.Width / 4.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Billboard Number", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 1 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Allocated", .ListView1.Width / 1 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Valid", .ListView1.Width / 1 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertSites WHERE  Discontinued = '" & "N" & "' AND Validstatus IS NULL AND  ApprovedStatus IS NOT NULL AND AllocationStatus IS NULL ORDER BY SiteNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))

    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(1) = CStr(rsLIST!SiteName)
    End If
    
    If Not IsNull(rsLIST!BBNo) Then
        MyList.SubItems(2) = CStr(rsLIST!BBNo)
    End If
    
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(3) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!sitephysicalAddress) Then
        MyList.SubItems(4) = CStr(rsLIST!sitephysicalAddress)
    End If
    
               
    If IsNull(rsLIST!AllocationStatus) Then
        MyList.SubItems(5) = CStr("NO")
    ElseIf Not IsNull(rsLIST!AllocationStatus) Then
        If rsLIST!AllocationStatus = "Y" Then
            MyList.SubItems(5) = CStr("YES")
        Else
            MyList.SubItems(5) = CStr("NO")
        End If
    End If
    
    
    If IsNull(rsLIST!Validstatus) Then
        MyList.SubItems(6) = CStr("NO")
    ElseIf Not IsNull(rsLIST!Validstatus) Then
        If rsLIST!Validstatus = "Y" Then
            MyList.SubItems(6) = CStr("YES")
        Else
            MyList.SubItems(6) = CStr("NO")
        End If
    End If
    
    
    rsLIST.MoveNext
    
Wend

'.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
'.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub



Private Sub GetNonDiscontinuedAdverts()

'On Error GoTo Err
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Advert Code ", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Type", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Length", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Width", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Weight", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Duration", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Days", .ListView1.Width / 5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertBBDetails  WHERE Discontinued='" & "N" & "'  AND BillBoardNo IS NOT NULL  ORDER BY BillBoardNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!BillBoardNo))

    If Not IsNull(rsLIST!Name) Then
        MyList.SubItems(1) = CStr(rsLIST!Name)
    End If
    
    If Not IsNull(rsLIST!TypeName) Then
        MyList.SubItems(2) = Trim(CStr(rsLIST!TypeName))
    End If
    
    If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!Length))
    End If
    
    If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(4) = Trim(CStr(rsLIST!Width))
    End If
    
    If Not IsNull(rsLIST!Weight) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!Weight))
    End If
    
    If Not IsNull(rsLIST!DurationName) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!DurationName))
    End If
    
    If Not IsNull(rsLIST!NoOfDays) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!NoOfDays))
    End If
    
    
            
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub GetUnApprovedContractsStructure()
'On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Contract No", .ListView2.Width / 5
.ListView2.ColumnHeaders.Add , , "Client Code", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Client Name", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Person", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Title", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Physical Address", .ListView2.Width / 3.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Address", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "City", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Mobile Phone", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Phone", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "E-Mail", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Fax", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Begins", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Ends", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Total Cost", .ListView2.Width / 5.5 ', lvwColumnCenter


.ListView2.View = lvwReport
End With
Exit Sub
Err:
    ErrorMessage
End Sub


Private Sub GetUnApprovedContracts()
'On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Contract No", .ListView2.Width / 5
.ListView2.ColumnHeaders.Add , , "Client Code", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Client Name", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Person", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Title", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Physical Address", .ListView2.Width / 3.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Address", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "City", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Mobile Phone", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Phone", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "E-Mail", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Fax", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Begins", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Ends", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Total Cost", .ListView2.Width / 5.5 ', lvwColumnCenter


.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertClients A,AdvertContractRequisition B WHERE A.CustomerId = B.ClientCode AND B.ApprovedStatus IS NULL AND B.EndDate > '" & Today & "'ORDER BY B.PurchaseOrderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))

    If Not IsNull(rsLIST!ClientCode) Then
        MyList.SubItems(1) = CStr(rsLIST!ClientCode)
    End If
    
    If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(2) = CStr(rsLIST!ClientName)
    End If
    
    If Not IsNull(rsLIST!ContactPerson) Then
        MyList.SubItems(3) = CStr(rsLIST!ContactPerson)
    End If
    
    If Not IsNull(rsLIST!ContactTitle) Then
        MyList.SubItems(4) = CStr(rsLIST!ContactTitle)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!PhysicalAddress))
    End If
    
    If Not IsNull(rsLIST!Address) Then
        MyList.SubItems(6) = CStr(rsLIST!Address)
    End If
    
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(7) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!MobilePhone) Then
        MyList.SubItems(8) = CStr(rsLIST!MobilePhone)
    End If
    
    If Not IsNull(rsLIST!Phone) Then
        MyList.SubItems(9) = CStr(rsLIST!Phone)
    End If
    
    If Not IsNull(rsLIST!Email) Then
        MyList.SubItems(10) = CStr(rsLIST!Email)
    End If
    
    If Not IsNull(rsLIST!Fax) Then
        MyList.SubItems(11) = CStr(rsLIST!Fax)
    End If
    
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(12) = CStr(rsLIST!StartDate)
    End If
    
    If Not IsNull(rsLIST!EndDate) Then
        MyList.SubItems(13) = CStr(rsLIST!EndDate)
    End If
    
    If Not IsNull(rsLIST!TotalCost) Then
        MyList.SubItems(14) = CStr(rsLIST!TotalCost)
    End If
    
            
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:

If Err.Number = 3265 Then Resume Next

End Sub
Private Sub GetUnApprovedRenewalContracts()
'On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Contract Number", .ListView2.Width / 5
.ListView2.ColumnHeaders.Add , , "Client Code", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Client Name", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Person", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Title", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Physical Address", .ListView2.Width / 3.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Address", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "City", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Mobile Phone", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Phone", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "E-Mail", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Fax", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Begins", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Ends", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Total Cost", .ListView2.Width / 5.5 ', lvwColumnCenter


.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertClients A,AdvertContractRequisition B WHERE A.CustomerId = B.ClientCode AND B.RenewalApprovalStatus IS NULL AND CompletionStatus = '" & "Y" & "'  ORDER BY B.PurchaseOrderNo", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))

    If Not IsNull(rsLIST!ClientCode) Then
        MyList.SubItems(1) = CStr(rsLIST!ClientCode)
    End If
    
    If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(2) = CStr(rsLIST!ClientName)
    End If
    
    If Not IsNull(rsLIST!ContactPerson) Then
        MyList.SubItems(3) = CStr(rsLIST!ContactPerson)
    End If
    
    If Not IsNull(rsLIST!ContactTitle) Then
        MyList.SubItems(4) = CStr(rsLIST!ContactTitle)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!PhysicalAddress))
    End If
    
    If Not IsNull(rsLIST!Address) Then
        MyList.SubItems(6) = CStr(rsLIST!Address)
    End If
    
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(7) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!MobilePhone) Then
        MyList.SubItems(8) = CStr(rsLIST!MobilePhone)
    End If
    
    If Not IsNull(rsLIST!Phone) Then
        MyList.SubItems(9) = CStr(rsLIST!Phone)
    End If
    
    If Not IsNull(rsLIST!Email) Then
        MyList.SubItems(10) = CStr(rsLIST!Email)
    End If
    
    If Not IsNull(rsLIST!Fax) Then
        MyList.SubItems(11) = CStr(rsLIST!Fax)
    End If
    
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(12) = CStr(rsLIST!StartDate)
    End If
    
    If Not IsNull(rsLIST!EndDate) Then
        MyList.SubItems(13) = CStr(rsLIST!EndDate)
    End If
    
    If Not IsNull(rsLIST!TotalCost) Then
        MyList.SubItems(14) = CStr(rsLIST!TotalCost)
    End If
    
            
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:

If Err.Number = 3265 Then Resume Next

End Sub

Private Sub GetUnApprovedSites()
'On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Site Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site City", .ListView1.Width / 4.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site Physical Address", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Rent Fee", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Rent Interval", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Council Fee", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Council Interval", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Finish Date", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Landlord", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Landlord Physical Address", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Landlord City", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Telephone", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Mobile No", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "E-Mail", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Fax", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Landlord No", .ListView1.Width / 5.5 ', lvwColumnCenter


.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM Advertsites A,AdvertSiteLords B WHERE A.LandLordNo = B.LandLordNo AND A.ApprovedStatus IS NULL AND A.ContractFinish > '" & Today & "' ORDER BY A.SiteNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))

    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(1) = CStr(rsLIST!SiteName)
    End If
    
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(2) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!sitephysicalAddress) Then
        MyList.SubItems(3) = CStr(rsLIST!sitephysicalAddress)
    End If
    
    If Not IsNull(rsLIST!SiteCharges) Then
        MyList.SubItems(4) = CStr(rsLIST!SiteCharges)
    End If
    
    If Not IsNull(rsLIST!SiteChargesInterval) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!SiteChargesInterval))
    End If
    
    If Not IsNull(rsLIST!CouncilCharges) Then
        MyList.SubItems(6) = CStr(rsLIST!CouncilCharges)
    End If
    
    If Not IsNull(rsLIST!CouncilChargesInterval) Then
        MyList.SubItems(7) = CStr(rsLIST!CouncilChargesInterval)
    End If
    
    If Not IsNull(rsLIST!ContractStart) Then
        MyList.SubItems(8) = CStr(rsLIST!ContractStart)
    End If
    
    If Not IsNull(rsLIST!ContractFinish) Then
        MyList.SubItems(9) = CStr(rsLIST!ContractFinish)
    End If
    
    If Not IsNull(rsLIST!Surname) Then
        MyList.SubItems(10) = CStr(rsLIST!Surname)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(11) = CStr(rsLIST!PhysicalAddress)
    End If
    
    If Not IsNull(rsLIST!LordCity) Then
        MyList.SubItems(12) = CStr(rsLIST!LordCity)
    End If
    
    If Not IsNull(rsLIST!TelephoneNo) Then
        MyList.SubItems(13) = CStr(rsLIST!TelephoneNo)
    End If
    
    If Not IsNull(rsLIST!MobileNo) Then
        MyList.SubItems(14) = CStr(rsLIST!MobileNo)
    End If
    
    If Not IsNull(rsLIST!Email) Then
        MyList.SubItems(15) = CStr(rsLIST!Email)
    End If
    
    If Not IsNull(rsLIST!Fax) Then
        MyList.SubItems(16) = CStr(rsLIST!Fax)
    End If
    
    If Not IsNull(rsLIST!LandLordNo) Then
        MyList.SubItems(17) = CStr(rsLIST!LandLordNo)
    End If
    
            
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:

If Err.Number = 3265 Then Resume Next

End Sub
Private Sub GetRenewalApprovalSites()
'On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Site Number", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Site Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site City", .ListView1.Width / 4.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site Physical Address", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Rent Fee", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Rent Interval", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Council Fee", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Council Interval", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Finish Date", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Landlord", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Landlord Physical Address", .ListView1.Width / 4.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Landlord City", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Telephone", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Mobile No", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "E-Mail", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Fax", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Landlord No", .ListView1.Width / 5.5 ', lvwColumnCenter


.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM Advertsites A,AdvertSiteLords B WHERE A.LandLordNo = B.LandLordNo AND A.RenewalApprovalStatus IS NULL  ORDER BY A.SiteNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))

    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(1) = CStr(rsLIST!SiteName)
    End If
    
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(2) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!sitephysicalAddress) Then
        MyList.SubItems(3) = CStr(rsLIST!sitephysicalAddress)
    End If
    
    If Not IsNull(rsLIST!SiteCharges) Then
        MyList.SubItems(4) = CStr(rsLIST!SiteCharges)
    End If
    
    If Not IsNull(rsLIST!SiteChargesInterval) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!SiteChargesInterval))
    End If
    
    If Not IsNull(rsLIST!CouncilCharges) Then
        MyList.SubItems(6) = CStr(rsLIST!CouncilCharges)
    End If
    
    If Not IsNull(rsLIST!CouncilChargesInterval) Then
        MyList.SubItems(7) = CStr(rsLIST!CouncilChargesInterval)
    End If
    
    If Not IsNull(rsLIST!ContractStart) Then
        MyList.SubItems(8) = CStr(rsLIST!ContractStart)
    End If
    
    If Not IsNull(rsLIST!ContractFinish) Then
        MyList.SubItems(9) = CStr(rsLIST!ContractFinish)
    End If
    
    If Not IsNull(rsLIST!Surname) Then
        MyList.SubItems(10) = CStr(rsLIST!Surname)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(11) = CStr(rsLIST!PhysicalAddress)
    End If
    
    If Not IsNull(rsLIST!LordCity) Then
        MyList.SubItems(12) = CStr(rsLIST!LordCity)
    End If
    
    If Not IsNull(rsLIST!TelephoneNo) Then
        MyList.SubItems(13) = CStr(rsLIST!TelephoneNo)
    End If
    
    If Not IsNull(rsLIST!MobileNo) Then
        MyList.SubItems(14) = CStr(rsLIST!MobileNo)
    End If
    
    If Not IsNull(rsLIST!Email) Then
        MyList.SubItems(15) = CStr(rsLIST!Email)
    End If
    
    If Not IsNull(rsLIST!Fax) Then
        MyList.SubItems(16) = CStr(rsLIST!Fax)
    End If
    
    If Not IsNull(rsLIST!LandLordNo) Then
        MyList.SubItems(17) = CStr(rsLIST!LandLordNo)
    End If
    
            
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:

If Err.Number = 3265 Then Resume Next

End Sub
 
 Sub chkSelectAll_Click()
'On Error GoTo Err
With Me
Dim i, j, k
j = .ListView1.ListItems.Count

If j = 0 Or .ListView1.View <> lvwReport Then .chkSelectAll.Value = 0: Exit Sub

Select Case .chkSelectAll.Value
Case 0
    For i = 1 To j
        .ListView1.ListItems(i).Checked = False
    Next i
Case 1
    For i = 1 To j
        .ListView1.ListItems(i).Checked = True
    Next i
Case Else
    Exit Sub
End Select
End With
Exit Sub
Err:
    ErrorMessage
End Sub
