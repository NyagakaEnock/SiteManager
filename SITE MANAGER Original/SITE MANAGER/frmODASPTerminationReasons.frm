VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPTerminationReasons 
   Caption         =   "Termination Reasons"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   9225
   Icon            =   "frmODASPTerminationReasons.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASPTerminationReasons.frx":0442
   ScaleHeight     =   4680
   ScaleWidth      =   9225
   Begin VB.Frame Frame12 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   9015
      Begin VB.CheckBox chkActive 
         Caption         =   "Active?"
         Height          =   375
         Left            =   8040
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtTerminationCode 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtTerminationReason 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   4815
      End
      Begin VB.Frame Frame4 
         Caption         =   "Operations"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   8775
         Begin VB.OptionButton optLandLord 
            Caption         =   "LandLord?"
            Height          =   255
            Left            =   3240
            TabIndex        =   16
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton optCompany 
            Caption         =   "Company?"
            Height          =   255
            Left            =   600
            TabIndex        =   15
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton optOthers 
            Caption         =   "Others?"
            Height          =   255
            Left            =   6480
            TabIndex        =   14
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   3015
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   7575
         Begin MSComctlLib.ListView ListView1 
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
               Name            =   "Arial"
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
            Picture         =   "frmODASPTerminationReasons.frx":0784
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
            Picture         =   "frmODASPTerminationReasons.frx":0886
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
            Picture         =   "frmODASPTerminationReasons.frx":0988
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
            Picture         =   "frmODASPTerminationReasons.frx":0A8A
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
            Picture         =   "frmODASPTerminationReasons.frx":0B8C
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
            Picture         =   "frmODASPTerminationReasons.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2520
            Width           =   855
         End
      End
      Begin VB.Label lblRelationshipCode 
         Caption         =   "Termination Code"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   435
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmODASPTerminationReasons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub loadRECORD()
On Error GoTo Err
    
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPTerminationReasons Where TerminationCode = '" & frmODASPTerminationReasons.txtTerminationCode.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    
    With rsCONTROL
        
        If .EOF Or .EOF Then Exit Sub
        
        frmODASPTerminationReasons.txtTerminationCode = !TerminationCode
        frmODASPTerminationReasons.txtTerminationReason = !TerminationReason
                             
        If !Status = "A" Then
                frmODASPTerminationReasons.chkActive.Value = 1
        Else: frmODASPTerminationReasons.chkActive.Value = 0
        End If
        
        If !LandLord = "Y" Then
                frmODASPTerminationReasons.optLandLord.Value = True
        Else: frmODASPTerminationReasons.optLandLord.Value = False
        End If
        
        If !Company = "Y" Then
                frmODASPTerminationReasons.optCompany.Value = True
        Else: frmODASPTerminationReasons.optCompany.Value = False
        End If
        
        If !Others = "Y" Then
                frmODASPTerminationReasons.optOthers.Value = True
        Else: frmODASPTerminationReasons.optOthers.Value = False
        End If
    
                
    End With

Exit Sub

Err:
    ErrorMessage
End Sub


Private Sub cmdAddNew_Click()
        baddRECORD = True
        clearALLRECORD
        enableALLRECORD
'        disableButtons
        frmODASPTerminationReasons.chkActive.Value = 1
End Sub


Private Sub cmdCancel_Click()
'        enableButtons
        clearALLRECORD
        disableALLRECORD
        baddRECORD = False
End Sub


Private Sub cmdDelete_Click()
On Error GoTo Err

If txtTerminationCode.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
'            Set rsCONTROL = New ADODB.Recordset
'
'            strSQL = "Select * from ODASPTerminationReasons Where TerminationCode = '" & frmODASPTerminationReasons.txtTerminationCode.Text & "'"
'            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsCONTROL = cnCOMMON.Execute("DELETE FROM ODASPTerminationReasons WHERE TerminationCode = '" & frmODASPTerminationReasons.txtTerminationCode.Text & "'")
            
'            With rsCONTROL
'
'                If .EOF And .BOF Then Exit Sub
'                .Delete
'                .Requery
                clearALLRECORD
                showALLTerminationReasons
'            End With
    End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
Exit Sub

Err:
    ErrorMessage

End Sub

Private Sub cmdEdit_Click()
        editMYRECORD
End Sub

Private Sub ValidateRECORD()
On Error GoTo Err

        bSaveRECORD = False
        
        If Screen.ActiveForm.txtTerminationCode.Text = "" Then
                MsgBox "The Operation Type MUST be Entered"
                Screen.ActiveForm.txtTerminationCode.SetFocus
        ElseIf Screen.ActiveForm.txtTerminationReason.Text <= "" Then
                MsgBox "The Description of the Operation cannot be Left Blank"
                txtTerminationReason.SetFocus
        Else
                bSaveRECORD = True
        End If
Exit Sub

Err:
    ErrorMessage
End Sub

Private Sub SaveRECORD()
On Error GoTo Err
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPTerminationrEASONS Where TerminationCode = '" & frmODASPTerminationReasons.txtTerminationCode.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
   With rsCONTROL
        If .BOF Or .EOF Then
                .AddNew
                !TerminationCode = frmODASPTerminationReasons.txtTerminationCode
                !PreparedBy = CurrentUserName
                !DatePrepared = Date
        End If
        
        !TerminationReason = frmODASPTerminationReasons.txtTerminationReason
        
        If frmODASPTerminationReasons.chkActive = 1 Then
                !Status = "A"
            Else: !Status = "I"
        End If

        If frmODASPTerminationReasons.optCompany = True Then
                !Company = "Y"
            Else: !Company = "N"
        End If
        
        If frmODASPTerminationReasons.optLandLord = True Then
                !LandLord = "Y"
        Else: !LandLord = "N"
        End If
        
        If frmODASPTerminationReasons.optOthers = True Then
                !Others = "Y"
        Else: !Others = "N"
        End If
       
        bSaveRECORD = False
        
         .Update
         .Requery
  End With
Exit Sub

Err:
    If Err.Number = -2147217873 Or Err.Number = -2147467259 Or Err.Number = -2147352571 Then
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
'                    enableButtons
                    disableALLRECORD
                    baddRECORD = False
                End If
        End If
        showALLTerminationReasons
        
End Sub

Private Sub cmdSearch_Click()
        searchMyRecord
End Sub

Private Sub Form_Activate()
    disableALLRECORD
'    enableButtons
    
    showALLTerminationReasons
End Sub

Private Sub Form_Load()

    OpenODBCConnection
      
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo Err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPTerminationReasons.txtTerminationCode.Text = Item.Text
            loadRECORD
        Else
            Item.Checked = False
        End If
        

Exit Sub

Err:
    ErrorMessage
End Sub



