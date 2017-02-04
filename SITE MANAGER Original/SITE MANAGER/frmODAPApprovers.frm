VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODAPApprovers 
   Caption         =   "Loan Approvers"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   11340
   Icon            =   "frmODAPApprovers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   11340
   Begin VB.Frame Frame12 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   10080
         TabIndex        =   26
         Top             =   2640
         Width           =   1095
         Begin VB.CommandButton cmdAdd 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODAPApprovers.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODAPApprovers.frx":040C
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   990
            Width           =   855
         End
         Begin VB.CommandButton cmdSearch 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODAPApprovers.frx":050E
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1365
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODAPApprovers.frx":0610
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1740
            Width           =   855
         End
         Begin VB.CommandButton cmdCancel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODAPApprovers.frx":0712
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2115
            Width           =   855
         End
         Begin VB.CommandButton cmdPrint 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODAPApprovers.frx":0814
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   2520
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   3855
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   4815
         Begin MSComctlLib.ListView ListView1 
            Height          =   3495
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   6165
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   11055
         Begin VB.TextBox txtOperationDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4560
            TabIndex        =   21
            Top             =   240
            Width           =   5775
         End
         Begin VB.ComboBox cboOperationType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   3
            Top             =   255
            Width           =   3375
         End
         Begin VB.Label Label5 
            Caption         =   "Operation "
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   330
            Width           =   1215
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Properties"
         Height          =   1815
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   11055
         Begin VB.TextBox txtReenterPassword 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9120
            PasswordChar    =   "x"
            TabIndex        =   5
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtLimitAmount 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   10
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtDateRetired 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5520
            TabIndex        =   9
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtPassword 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   5520
            PasswordChar    =   "x"
            TabIndex        =   2
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtDateAssigned 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9120
            TabIndex        =   8
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtUserCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   4
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtNames 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4560
            TabIndex        =   6
            Top             =   240
            Width           =   6015
         End
         Begin VB.ComboBox cboStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   7
            Top             =   1215
            Width           =   2535
         End
         Begin VB.Label Label6 
            Caption         =   "Re-Enter Password"
            Height          =   195
            Left            =   7320
            TabIndex        =   24
            Top             =   810
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Limit"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Date Retired"
            Height          =   195
            Left            =   4560
            TabIndex        =   18
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Left            =   4560
            TabIndex        =   17
            Top             =   810
            Width           =   855
         End
         Begin VB.Label lblBenefitCode 
            Caption         =   "User Code"
            Height          =   315
            Left            =   240
            TabIndex        =   16
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Name"
            Height          =   195
            Left            =   3960
            TabIndex        =   15
            Top             =   330
            Width           =   855
         End
         Begin VB.Label lblStatus 
            Caption         =   "Status"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   1290
            Width           =   735
         End
         Begin VB.Label lblDateAssigned 
            Caption         =   "Date Assigned"
            Height          =   195
            Left            =   7320
            TabIndex        =   13
            Top             =   1260
            Width           =   1215
         End
      End
      Begin VB.Frame Frame14 
         Height          =   3855
         Left            =   5040
         TabIndex        =   11
         Top             =   2640
         Width           =   4935
         Begin MSComctlLib.ListView ListView2 
            Height          =   3495
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   6165
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
End
Attribute VB_Name = "frmODAPApprovers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAPPROVERS As clsODAPApprovers
Public rsLOAN As ADODB.Recordset

Private Sub cboOperationType_GotFocus()
        Set rsAPPROVERS = New cLoanApprovers
        rsAPPROVERS.operationtypeGOTFOCUS
        Set rsAPPROVERS = Nothing

End Sub
Private Sub cboOperationType_KeyPress(Index As Integer)
On Error GoTo err

        Set rsAPPROVERS = New cLoanApprovers
        Set rsAPPROVERS = Nothing

Exit Sub

err:
    ErrorMessage

End Sub
Private Sub cboOperationType_LostFocus()
        Set rsAPPROVERS = New cLoanApprovers
        rsAPPROVERS.operationtypeLOSTFOCUS
        Set rsAPPROVERS = Nothing
        GetUserCode

End Sub
Private Sub cboStatus_gotFocus()
        Set rsAPPROVERS = New cLoanApprovers
        rsAPPROVERS.statusGOTFOCUS
        Set rsAPPROVERS = Nothing
End Sub
Private Sub cbostatus_keypress(Index As Integer)
On Error GoTo err
        KeyAscii = 0
Exit Sub

err:
    ErrorMessage

End Sub
Private Sub cboStatus_LostFocus()
On Error GoTo err

        Set rsAPPROVERS = New cLoanApprovers
        rsAPPROVERS.statusLOSTFOCUS
        Set rsAPPROVERS = Nothing

Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cmdAdd_Click()
        
        Set rsAPPROVERS = New cLoanApprovers
        rsAPPROVERS.addRECORD
        Set rsAPPROVERS = Nothing
End Sub

Private Sub ClearListview()
On Error GoTo err

 Dim j, i As Integer
       
                j = ListView1.ListItems.Count
            
        For i = 1 To j
                ListView1.ListItems(i).Checked = False
        Next i
Exit Sub
err:
ErrorMessage
    
End Sub

Private Sub cmdCancel_Click()
        Set rsAPPROVERS = New cLoanApprovers
        rsAPPROVERS.Cancelrecord
        Set rsAPPROVERS = Nothing
End Sub

Private Sub cmdDelete_Click()
On Error GoTo err

        Set rsAPPROVERS = New cLoanApprovers
        rsAPPROVERS.deleteRECORD
        Set rsAPPROVERS = Nothing

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdEdit_Click()
    EditMyRecord
End Sub


Private Sub cmdSearch_Click()
        searchMyRecord
End Sub


Private Sub cmdUpdate_Click()
        Set rsAPPROVERS = New cLoanApprovers
        rsAPPROVERS.updateRECORD
        Set rsAPPROVERS = Nothing
End Sub

Private Sub Form_Activate()
        enableButtons
        disableALLRECORD
        Set rsAPPROVERS = New cLoanApprovers
        rsAPPROVERS.DisplayOperationType
        rsAPPROVERS.operationtypeLOSTFOCUS
        Set rsAPPROVERS = Nothing
        GetUserCode
        GetAPPROVERS

End Sub

'Private Sub GetCheckedBoxes()
'On Error GoTo err
'
'            Dim j, i As Integer, strRequirementcode As String
'
'            j = ListView1.ListItems.Count
'
'            For i = 1 To j
'                    If ListView1.ListItems(i).Checked = True Then
'                        strRequirementcode = ListView1.ListItems(i).Text
'                        frmALISPLoanApprover.txtUserCode.Text = strRequirementcode
'
'                        UpdateALLRECORDS
'                    End If
'                    strRequirementcode = ""
'            Next i
'
'
'Exit Sub
'
'err:
'    ErrorMessage
'End Sub
Private Sub Form_Load()
On Error GoTo err
  
        Call OpenConnection
        
        'create the instance of the data source class
        'Set rsAPPROVERS = New cLoanApprovers
        'Call rsAPPROVERS.disableRECORD
        
       
        'Set rsLOAN = New Recordset
        'strLOAN = "SELECT * from ALISPLoanApprover;"

        'rsLOAN.Open strLOAN, cnCOMMON, adOpenKeyset, adLockOptimistic

        'rsAPPROVERS.LoadGrid

        
        
        
Exit Sub
err:

ErrorMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
        baddRECORD = False
        bsearchRECORD = False
        beditRECORD = False
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
                        
            frmALISPLoanApprover.txtUserCode.Text = Item.Text
            loadDETAILS
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView2.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                End If
            Next i
                        
            frmALISPLoanApprover.txtUserCode.Text = Item.Text
            Set rsAPPROVERS = New cLoanApprovers
            rsAPPROVERS.searchRECORD
            Set rsAPPROVERS = Nothing
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub loadDETAILS()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset

        strSQL = "Select * from AdminUserRegister where StaffIDNO = '" & frmALISPLoanApprover.txtUserCode & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .BOF Or .EOF Then Exit Sub
                frmALISPLoanApprover.txtDateAssigned.Text = Date
                frmALISPLoanApprover.txtLimitAmount.Text = 0
                frmALISPLoanApprover.txtNames.Text = !UserName
                frmALISPLoanApprover.txtUserCode.Text = !StaffIdNo
                frmALISPLoanApprover.cboStatus = "A"
        End With
        
        rsCONTROL.Close
        strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub txtDateAssigned_Click()
    frmALISPLoanApprover.txtDateAssigned.Text = Date
    frmALISPLoanApprover.txtDateRetired.Text = DateAdd("d", 20, frmALISPLoanApprover.txtDateAssigned.Text)
End Sub

Private Sub txtReenterPassword_LostFocus()
        With frmALISPLoanApprover
            If Trim(.txtPassword.Text) <> Trim(.txtReenterPassword.Text) Then
                MsgBox "The Password you entered do not Match", vbOKOnly
            End If
        End With
End Sub
