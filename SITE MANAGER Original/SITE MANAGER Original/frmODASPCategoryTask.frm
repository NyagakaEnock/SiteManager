VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPCategoryTask 
   Caption         =   "Media Category"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   9225
   Icon            =   "frmODASPCategoryTask.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASPCategoryTask.frx":0442
   ScaleHeight     =   4740
   ScaleWidth      =   9225
   Begin VB.Frame Frame12 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   9135
      Begin VB.CheckBox chkActive 
         Caption         =   "Active?"
         Height          =   375
         Left            =   7680
         TabIndex        =   19
         Top             =   810
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Actual Tasks Set for the Category"
         Height          =   3015
         Left            =   3960
         TabIndex        =   17
         Top             =   1440
         Width           =   3735
         Begin MSComctlLib.ListView ListView2 
            Height          =   2655
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.TextBox txtDepartmentDescription 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2760
         TabIndex        =   15
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox txtDeptCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtCategoryCode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtCategoryDescription 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   5655
      End
      Begin VB.Frame Frame1 
         Caption         =   "List All Departments/Tasks"
         Height          =   3015
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   3855
         Begin MSComctlLib.ListView ListView1 
            Height          =   2655
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
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
               Size            =   9
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
         TabIndex        =   4
         Top             =   1440
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
            Picture         =   "frmODASPCategoryTask.frx":0784
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
            Picture         =   "frmODASPCategoryTask.frx":0886
            Style           =   1  'Graphical
            TabIndex        =   7
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
            TabIndex        =   10
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
            Picture         =   "frmODASPCategoryTask.frx":0988
            Style           =   1  'Graphical
            TabIndex        =   9
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
            Picture         =   "frmODASPCategoryTask.frx":0A8A
            Style           =   1  'Graphical
            TabIndex        =   8
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
            Picture         =   "frmODASPCategoryTask.frx":0B8C
            Style           =   1  'Graphical
            TabIndex        =   6
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
            Picture         =   "frmODASPCategoryTask.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2520
            Width           =   855
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Dept"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   915
         Width           =   1335
      End
      Begin VB.Label lblRelationshipCode 
         Caption         =   "Category Code"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   435
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmODASPCategoryTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub loadRECORD()
'On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPCategoryTask Where CategoryCode = '" & frmODASPCategoryTask.txtCategoryCode.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    
    With rsCONTROL
        
        If .EOF Or .EOF Then Exit Sub
        
        frmODASPCategoryTask.txtCategoryCode = !categorycode
        frmODASPCategoryTask.txtCategoryDescription = !categoryname
                            
        If !Status = "A" Then
                frmODASPCategoryTask.chkActive.Value = 1
        Else: frmODASPCategoryTask.chkActive.Value = 0
        End If
    End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdAddNew_Click()
        baddRECORD = True
        enableALLRECORD
        disableButtons
End Sub


Private Sub cmdCancel_Click()
        enableButtons
        clearRELN
        disableRELN
        baddRECORD = False
End Sub


Private Sub cmdDelete_Click()
'On Error GoTo err

If txtCategoryCode.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            Set rsCONTROL = New ADODB.Recordset
    
            strSQL = "Select * from ODASPCategoryCode Where CategoryCode = '" & frmODASPCategoryTask.txtCategoryCode.Text & "'"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            
            With rsCONTROL
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearALLRECORD
                getALLOPERATIONS
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

        bSaveRECORD = False
        
        If Screen.ActiveForm.txtCategoryCode.Text = "" Then
                MsgBox "The Operation Type MUST be Entered"
                Screen.ActiveForm.txtCategoryCode.SetFocus
        ElseIf Screen.ActiveForm.txtCategoryDescription.Text <= "" Then
                MsgBox "The Description of the Category cannot be Left Blank"
                txtCategoryDescription.SetFocus
        Else
                bSaveRECORD = True
        End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveRECORD()
''On Error GoTo err
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPCategoryTask Where CategoryCode = '" & frmODASPCategoryTask.txtCategoryCode.Text & "' and departmentCode = '" & frmODASPCategoryTask.txtDeptCode.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
   With rsCONTROL
        If .BOF Or .EOF Then
                .AddNew
                !categorycode = frmODASPCategoryTask.txtCategoryCode
                !PreparedBy = CurrentUserName
                !DatePrepared = Date
        End If
        
        !DepartmentCode = frmODASPCategoryTask.txtDeptCode.Text
        
        If frmODASPCategoryTask.chkActive = 1 Then
                !Status = "A"
            Else: !Status = "I"
        End If

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
            SaveRECORD
                If bSaveRECORD = False Then
                    enableButtons
                    disableALLRECORD
                    baddRECORD = False
                End If
        End If
        showALLTASKS
        showALLDEPTS
End Sub

Private Sub cmdSearch_Click()
        searchMyRecord
End Sub

Private Sub Form_Activate()
    disableALLRECORD
    enableButtons
    showALLDEPTS
    showALLTASKS
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
                        
            frmODASPCategoryTask.txtDeptCode.Text = Item.Text
            frmODASPCategoryTask.txtDepartmentDescription.Text = Item.SubItems(1)
            showALLTASKS
            'loadRECORD
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
                        
            frmODASPCategoryTask.txtDeptCode.Text = Item.Text
            frmODASPCategoryTask.txtDepartmentDescription.Text = Item.SubItems(1)
            loadRECORD
         Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub


