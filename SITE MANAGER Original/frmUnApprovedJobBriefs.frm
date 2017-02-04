VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmUnApprovedJobBriefs 
   Caption         =   "Unapproved quotation items under current quotation No."
   ClientHeight    =   2205
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSelectAll 
      Caption         =   "Check1"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   9015
      Begin VB.CommandButton cmdApprove 
         BackColor       =   &H000000FF&
         Caption         =   "&APPROVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         MaskColor       =   &H008080FF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4800
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   2566
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuAuthorise 
      Caption         =   "Quotation Authorisations"
      Visible         =   0   'False
      Begin VB.Menu mnuApproveQuotation 
         Caption         =   "Approve Job Brief"
      End
      Begin VB.Menu mnuGJjnHGHJNJK 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAuthorizeQuota 
         Caption         =   "Authorise Job Brief"
      End
   End
End
Attribute VB_Name = "frmUnApprovedJobBriefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MediaCode
Public rsJOBBrief, rsCATEGORY As ADODB.Recordset

Private Sub Command1_Click()

End Sub

Private Sub chkSelectAll_Click()
''On Error GoTo Err
With Screen.ActiveForm
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
err:
    ErrorMessage
End Sub

Private Sub cmdApprove_Click()
''On Error GoTo Err
With Screen.ActiveForm
Select Case cmdApprove.Caption
  Case "&APPROVE"

If .txtPassword.Text = Empty Then
    MsgBox "Please enter password", vbExclamation, "Quotation Approval": Exit Sub
    Else
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM AdminIndividualRights WHERE Username = '" & CurrentUserName & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    
       If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
         If rsFindRecord!App = 1 Then GoTo ConfirmPassword
         If rsFindRecord!App = 0 Then
         MsgBox "You dont have right to approve quotations....Please consult your System Administrator", vbExclamation, "Approval Denied"
         End If: Exit Sub
       End If
          Set rsFindRecord = Nothing
       
  
ConfirmPassword:
        Set rsFindRecord = New ADODB.Recordset
          rsFindRecord.Open "SELECT * FROM Adminuserregister WHERE Username = '" & CurrentUserName & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
          
          If rsFindRecord.BOF And rsFindRecord.EOF Then Exit Sub
           If rsFindRecord!Password <> .txtPassword.Text Then
           MsgBox "The Password you entered is not correct...Please Re-enter the password OR consult your system administrator", vbExclamation, "Wrong Password"
           Else
             Call ApproveSelectedQuotationNo
              .Frame1.Visible = False
               Unload frmUnApprovedQuotationItems
               RemoveCurrentListItem
           End If
           
           
   Case "&AUTHORIZE"
   
   If .txtPassword.Text = Empty Then
    MsgBox "Please enter password", vbExclamation, "Quotation Approval": Exit Sub
    Else
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM AdminIndividualRights WHERE Username = '" & CurrentUserName & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    
       If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
         If rsFindRecord!AuthorisationRights = 1 Then GoTo ConfirmAuthorisationPassword
         If rsFindRecord!AuthorisationRights = 0 Then
         MsgBox "You dont have right to approve quotations....Please consult your System Administrator", vbExclamation, "Approval Denied"
         End If: Exit Sub
       End If
          Set rsFindRecord = Nothing
       
  
ConfirmAuthorisationPassword:
        Set rsFindRecord = New ADODB.Recordset
          rsFindRecord.Open "SELECT * FROM Adminuserregister WHERE Username = '" & CurrentUserName & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
          
          If rsFindRecord.BOF And rsFindRecord.EOF Then Exit Sub
           If rsFindRecord!Password <> .txtPassword.Text Then
           MsgBox "The Password you entered is not correct...Please Re-enter the password OR consult your system administrator", vbExclamation, "Wrong Password"
           Else
             Call AuthorizeSelectedQuotationNo
              .Frame1.Visible = False
               Unload frmUnApprovedQuotationItems
               RemoveCurrentListItem
           End If
        
   
   Case Else
   Exit Sub
 End Select
    
End With
Exit Sub
err:
   ErrorMessage
End Sub
Public Sub RemoveCurrentListItem()
''On Error GoTo Err
With SchedulingMain
Dim i, j, k
   j = .ListView1.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView1.ListItems(i).Checked = True Then
         .ListView1.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
err:
   ErrorMessage
End Sub
Private Sub ApproveSelectedQuotationNo()
''On Error GoTo Err
  Set rsLineUpdate = New ADODB.Recordset
    rsLineUpdate.Open ("UPDATE AdvertJobBrief SET ApprovedBy = '" & CurrentUserName & "',DateApproved = '" & MyCurrentDate & "',ApprovedStatus = '" & "Y" & "' WHERE JobBriefNo = '" & JobBriefNo & "'"), cnCOMMON, adOpenKeyset, adLockOptimistic
  Set rsLineUpdate = Nothing
  MsgBox "The Job brief number  " & JobBriefNo & " has successfully been approved", vbInformation, "Quotation Approval"
   
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub AuthorizeSelectedQuotationNo()
With Screen.ActiveForm

''On Error GoTo Err
  Set rsLineUpdate = New ADODB.Recordset
        rsLineUpdate.Open ("UPDATE AdvertJobBrief SET AuthorizedBy = '" & CurrentUserName & "',DateAuthorized = '" & MyCurrentDate & "',AuthorizationStatus = '" & "Y" & "' WHERE Jobbriefno = '" & JobBriefNo & "'"), cnCOMMON, adOpenKeyset, adLockOptimistic
  Set rsLineUpdate = Nothing
  MsgBox "The Job brief number  " & JobBriefNo & "  has successfully been AUTHORIZED", vbInformation, "Quotation Approval"
        
        Set rsJOBBrief = New ADODB.Recordset
        rsJOBBrief.Open "SELECT * FROM Advertjobbriefitems WHERE Jobbriefno = '" & JobBriefNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsJOBBrief.BOF And rsJOBBrief.EOF Then Exit Sub

        
        Do While Not rsJOBBrief.EOF
                
                Set rsCATEGORY = New ADODB.Recordset
                rsCATEGORY.Open "SELECT * FROM AdvertParamCategoryConfig WHERE CategoryCode = '" & rsJOBBrief!categorycode & "'", cnCOMMON, adOpenKeyset, adLockOptimistic

                If rsCATEGORY.BOF And rsCATEGORY.EOF Then Exit Sub

                Set rsNewRecord = New ADODB.Recordset
                     
                NewSQL = "INSERT INTO AdvertJobCard(JobCardNo,DeptCode,Status,TotalCost,AccPeriod,StartDate,DateCreated,CreatedBy)VALUES('" & JobBriefNo & "','" & rsCATEGORY!DeptCode & "','START'," & CCur(0) & ",'" & Trim(MyCurrentPeriod) & "','" & MyCurrentDate & "','" & MyCurrentDate & "','" & CurrentUserName & "');"
                
                rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Set rsNewRecord = Nothing

                rsJOBBrief.MoveNext
        Loop
            
      
    
    
   End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
    With Screen.ActiveForm
            ShowAllUnApprovedJobBriefs
            .chkSelectAll.Value = 1
             Call chkSelectAll_Click
                    
            .Frame1.Visible = False
            .Label2.Caption = "YOU ARE LOGGED IN AS '" & CurrentUserName & "'"
    End With
End Sub


Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Screen.ActiveForm
    If Button = 2 Then
       If SchedulingMain.ListView1.ColumnHeaders(1).Text = "Job Brief No" Then
        .mnuAuthorizeQuota.Enabled = False
        PopupMenu mnuAuthorise, , , , mnuApproveQuotation
        ElseIf SchedulingMain.ListView1.ColumnHeaders(1).Text = "JobBrief No " Then
        .mnuApproveQuotation.Enabled = False
        PopupMenu mnuAuthorise, , , , mnuAuthorizeQuota
        End If
    End If
End With
End Sub

Private Sub mnuApproveQuotation_Click()
        With Screen.ActiveForm
                .Frame1.Visible = True
        
        End With
End Sub

Private Sub mnuAuthorizequota_Click()
        With Screen.ActiveForm
        .Frame1.Visible = True
        .cmdApprove.Caption = "&AUTHORIZE"
        End With
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
''On Error GoTo Err
        If KeyAscii = vbKeyReturn Then
                Call cmdApprove_Click
        Else
        
        End If
Exit Sub

err:
   ErrorMessage
End Sub
