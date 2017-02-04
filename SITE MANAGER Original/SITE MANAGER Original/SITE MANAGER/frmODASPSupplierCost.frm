VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPSupplierCost 
   Caption         =   "Departmental Access"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   9225
   Icon            =   "frmODASPSupplierCost.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASPSupplierCost.frx":0442
   ScaleHeight     =   5640
   ScaleWidth      =   9225
   Begin VB.Frame Frame12 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9015
      Begin VB.Frame Frame3 
         Height          =   2295
         Left            =   4200
         TabIndex        =   13
         Top             =   120
         Width           =   4695
         Begin VB.TextBox txtNames 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1560
            TabIndex        =   19
            Top             =   1740
            Width           =   3015
         End
         Begin VB.TextBox txtStaffIDNo 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1560
            TabIndex        =   18
            Top             =   1230
            Width           =   1335
         End
         Begin VB.TextBox txtDepartmentDescription 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1560
            TabIndex        =   16
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtDepartmentCode 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1560
            TabIndex        =   15
            Top             =   210
            Width           =   3015
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Active?"
            Height          =   375
            Left            =   3000
            TabIndex        =   14
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Department"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Names"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1770
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Staff No"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label lblRelationshipCode 
            Caption         =   "Dept Code"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Users with Access"
         Height          =   3015
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   7575
         Begin MSComctlLib.ListView ListView2 
            Height          =   2655
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   4683
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
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
         Caption         =   "Users NOT in the Department"
         Height          =   2295
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   3975
         Begin MSComctlLib.ListView ListView1 
            Height          =   1935
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   3413
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
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
         Height          =   3015
         Left            =   7800
         TabIndex        =   2
         Top             =   2400
         Width           =   1095
         Begin VB.CommandButton cmdAddNew 
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
            Picture         =   "frmODASPSupplierCost.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   0
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
            Picture         =   "frmODASPSupplierCost.frx":0886
            Style           =   1  'Graphical
            TabIndex        =   5
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
            TabIndex        =   8
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
            Picture         =   "frmODASPSupplierCost.frx":0988
            Style           =   1  'Graphical
            TabIndex        =   7
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
            Picture         =   "frmODASPSupplierCost.frx":0A8A
            Style           =   1  'Graphical
            TabIndex        =   6
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
            Picture         =   "frmODASPSupplierCost.frx":0B8C
            Style           =   1  'Graphical
            TabIndex        =   4
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
            Picture         =   "frmODASPSupplierCost.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2520
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmODASPSupplierCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub loadRECORD()
'On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPDeptAccess, ODASPDepartment, AdminUserRegister Where AdminUserRegister.StaffIdNo = ODASPDeptAccess.StaffIdNo and ODASPDepartment.DepartmentCode = ODASPDeptAccess.DepartmentCode and ODASPDeptAccess.DepartmentCode = '" & frmODASPDeptAccess.txtDepartmentCode.Text & "' and ODASPDeptAccess.StaffIdNo = '" & frmODASPDeptAccess.txtStaffIDNo.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsCONTROL
        
        If .EOF Or .EOF Then Exit Sub
        
        frmODASPDeptAccess.txtDepartmentCode = !DepartmentCode
        frmODASPDeptAccess.txtDepartmentDescription = !DepartmentDescription
        frmODASPDeptAccess.txtStaffIDNo = !StaffIdNo
        frmODASPDeptAccess.txtNames = !AllNames
                     
        If !Status = "A" Then
                frmODASPDeptAccess.chkStatus.Value = 1
        Else: frmODASPDeptAccess.chkStatus.Value = 0
        End If
        
         
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub loadSTAFF()
'On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from AdminUserRegister Where AdminUserRegister.StaffIdNo = '" & frmODASPDeptAccess.txtStaffIDNo.Text & "' "
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsCONTROL
        
        If .EOF Or .EOF Then Exit Sub
        
        frmODASPDeptAccess.txtNames = !AllNames
                     
    End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdAddNew_Click()
        baddRECORD = True
        clearRECORD
        enableRECORD
        disableButtons
End Sub
Private Sub clearRECORD()
On Error GoTo err

    With frmODASPDeptAccess
        .txtNames.Text = Empty
        .txtStaffIDNo.Text = Empty
        .chkStatus.Value = 1
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub enableRECORD()
On Error GoTo err

    With frmODASPDeptAccess
        .txtNames.Locked = False
        .txtStaffIDNo.Locked = False
        .chkStatus.Value = 1
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdCancel_Click()
        enableButtons
        clearALLRECORD
        disableALLRECORD
        baddRECORD = False
End Sub


Private Sub cmdDelete_Click()
'On Error GoTo err

If txtDepartmentCode.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            Set rsCONTROL = New ADODB.Recordset
    
            strSQL = "Select * from ODASPDepartment Where DepartmentCode = '" & frmODASPDeptAccess.txtDepartmentCode.Text & "'"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            
            With rsCONTROL
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearRECORD
                showDepartmentACCESS
                showStaffACCESS
            End With
    End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cmdEdit_Click()
        editMYRECORD
End Sub

Private Sub ValidateRECORD()
'On Error GoTo err
    With frmODASPDeptAccess
        bSaveRECORD = False
        
        If .txtDepartmentCode.Text = Empty Then
                MsgBox "The Operation Type MUST be Entered"
                .txtDepartmentCode.SetFocus
        ElseIf .txtStaffIDNo.Text <= "" Then
                MsgBox "The Staff ID No is Required"
                .txtStaffIDNo.SetFocus
        Else
                bSaveRECORD = True
        End If
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub saveRECORD()
'On Error GoTo err
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPDeptAccess Where DepartmentCode = '" & frmODASPDeptAccess.txtDepartmentCode.Text & "' and StaffIDNo = '" & frmODASPDeptAccess.txtStaffIDNo.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
   With rsCONTROL
        If .BOF Or .EOF Then
                .AddNew
                !DepartmentCode = frmODASPDeptAccess.txtDepartmentCode
                !StaffIdNo = frmODASPDeptAccess.txtStaffIDNo.Text
                !PreparedBy = CurrentUserName
                !DatePrepared = Date
        End If
        
        
        If frmODASPDeptAccess.chkStatus.Value = 1 Then
                !Status = "A"
        Else: !Status = "I"
        End If
        
        !StatusDate = Date

        bSaveRECORD = False
        
         .Update
         .Requery
  End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsCONTROL.CancelUpdate
            rsCONTROL.Requery
    Else
        UpdateErrorMessage
    End If

End Sub


Private Sub cmdUpdate_Click()
        bSaveRECORD = True
        ValidateRECORD
        If bSaveRECORD = True Then
            saveRECORD
                If bSaveRECORD = False Then
                    enableButtons
                    disableALLRECORD
                    baddRECORD = False
                End If
        End If
        showDepartmentACCESS
        showStaffACCESS
End Sub

Private Sub cmdSearch_Click()
        searchMyRecord
End Sub

Private Sub Form_Activate()
    disableALLRECORD
    enableButtons
    showStaffACCESS
    showDepartmentACCESS
End Sub

Private Sub Form_Load()

    OpenODBCConnection
      
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPDeptAccess.txtStaffIDNo.Text = Item.Text
            loadSTAFF
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView2.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPDeptAccess.txtDepartmentCode.Text = Item.Text
            loadRECORD
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

