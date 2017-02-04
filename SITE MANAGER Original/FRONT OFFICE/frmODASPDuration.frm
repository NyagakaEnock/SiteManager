VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPDuration 
   Caption         =   "Duration Types"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   8880
   Icon            =   "frmODASPDuration.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASPDuration.frx":0442
   ScaleHeight     =   5430
   ScaleWidth      =   8880
   Begin VB.Frame Frame12 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
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
         Left            =   7560
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
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
            Height          =   495
            Left            =   120
            Picture         =   "frmODASPDuration.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   720
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
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1200
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
            Height          =   495
            Left            =   120
            Picture         =   "frmODASPDuration.frx":0886
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1680
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
            Height          =   495
            Left            =   120
            Picture         =   "frmODASPDuration.frx":0988
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2160
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
            Height          =   495
            Left            =   120
            Picture         =   "frmODASPDuration.frx":0A8A
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2640
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
            Height          =   495
            Left            =   120
            Picture         =   "frmODASPDuration.frx":0B8C
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3120
            Width           =   855
         End
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
            Height          =   495
            Left            =   120
            Picture         =   "frmODASPDuration.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame13 
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   8535
         Begin VB.CheckBox chkStatus 
            Caption         =   "Status?"
            Height          =   375
            Left            =   4680
            TabIndex        =   19
            Top             =   720
            Width           =   2175
         End
         Begin VB.Frame Frame1 
            Height          =   615
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   4215
            Begin VB.OptionButton optWeek 
               Caption         =   "Week?"
               Height          =   255
               Left            =   2040
               TabIndex        =   18
               Top             =   240
               Width           =   2055
            End
            Begin VB.OptionButton optDay 
               Caption         =   "Day?"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.TextBox txtDurationDescription 
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
            Left            =   3720
            TabIndex        =   4
            Top             =   240
            Width           =   4695
         End
         Begin VB.TextBox txtDurationMode 
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
            Left            =   960
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Description "
            Height          =   255
            Left            =   2760
            TabIndex        =   6
            Top             =   293
            Width           =   1095
         End
         Begin VB.Label lblRelationshipCode 
            Caption         =   "Mode"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   293
            Width           =   1335
         End
      End
      Begin VB.Frame Frame14 
         Height          =   3855
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   7335
         Begin MSComctlLib.ListView ListView1 
            Height          =   3495
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   6165
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   0
         End
      End
   End
End
Attribute VB_Name = "frmODASPDuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRELN As ADODB.Recordset, strRELN As String, bsaveRECORD As Boolean

Sub clearRELN()
        txtDurationMode.Text = ""
        txtDurationDescription.Text = ""
End Sub




Private Sub cmdAddNew_Click()
        ListView1.Enabled = False
        enableALLRECORD
        clearALLRECORD
        disableButtons
End Sub

Private Sub cmdCancel_Click()
        
        ListView1.Enabled = False
        enableButtons
        clearRELN
        disableALLRECORD
End Sub


Private Sub cmdDelete_Click()
On Error GoTo err

If txtDurationMode.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            Set rsSAVE = New Recordset
            strRELN = "SELECT * from ODASPDuration  where DurationMode = '" & frmODASPDuration.txtDurationMode & "';"
        
            rsSAVE.Open strRELN, cnCOMMON, adOpenKeyset, adLockOptimistic
        
            If rsSAVE.BOF Or rsSAVE.EOF Then Exit Sub

            With rsSAVE
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearALLRECORD
                showDURATIONMODE
                                         
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
On Error GoTo err
        If ListView1.Enabled = False Then
            ListView1.Enabled = True
            ListView1.SetFocus
            Exit Sub
        End If
        
            
            
'Dim strQRE As Variant
Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset

        Select Case cmdEdit.Caption
                Case "&Edit"
                        enableALLRECORD

                        'strQRE = InputBox("Enter The DurationMode to search.", "Search Value")
    
                        rsFind.Open "SELECT * FROM ODASPDuration WHERE DurationMode LIKE '" & SelectedListItem & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsFind
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                Else
                                        txtDurationMode = !DurationMode
                                        Me.txtDurationDescription = !DurationDescription
                                        Edit = True
                                End If
                        End With
        
                        If Edit Then
                                cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ODASPDuration WHERE DurationMode = '" & txtDurationMode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                            !DurationMode = txtDurationMode
                            !Duration = txtDurationDescription
                            .Update
                            .Requery
                            Edit = False
                    End With
                
                    cmdEdit.Caption = "&Edit"
            Case Else
        
            Exit Sub

        End Select

Exit Sub

err:

    If err.Number = 40009 Then
            MsgBox "Record requested does not exist in the Database! Check your Entries.", vbInformation, "Searching."
                rsFind.Requery

            If rsFind.BOF Then Exit Sub
                rsFind.MoveFirst

    ElseIf err.Number = 3021 Then
            MsgBox "Requested record not found! Refresh the database and try the search again...or Check your entries.", vbInformation, "Searching."
                rsFind.Requery

            If rsFind.BOF Then Exit Sub
                rsFind.MoveFirst
    Else
                ErrorMessage
End If

End Sub



Private Sub ValidateRELN()
On Error GoTo err
        With frmODASPDuration
                If .txtDurationDescription.Text = Empty Then
                        MsgBox "The Duration Code  is Required"
                        .txtDurationDescription.SetFocus
                
                ElseIf .txtDurationMode.Text <= Empty Then
                        MsgBox "The Duration Mode is  required"
                        .txtDurationMode.SetFocus
                
                Else
                        bsaveRECORD = True
                End If
        End With
        
Exit Sub

err:
ErrorMessage
End Sub

Private Sub saveRecord()
On Error GoTo err
    With frmODASPDuration
    
            Set rsSAVE = New Recordset
            strRELN = "SELECT * from ODASPDuration  where DurationMode = '" & .txtDurationMode & "';"
        
            rsSAVE.Open strRELN, cnCOMMON, adOpenKeyset, adLockOptimistic
        
            If rsSAVE.BOF Or rsSAVE.EOF Then
                    rsSAVE.AddNew
                    rsSAVE!DurationMode = .txtDurationMode
                    rsSAVE!DurationDescription = txtDurationDescription
                    rsSAVE!Preparedby = CurrentUserName
                    rsSAVE!dateprepared = Date
            End If
            
            If .chkstatus.Value = 1 Then
                    rsSAVE!Status = "A"
                    rsSAVE!StatusDate = Date
            Else: rsSAVE!Status = "I"
                    rsSAVE!StatusDate = Date
            End If
                        
            If .optDay.Value = True Then
                    rsSAVE!DayType = "Y"
            Else: rsSAVE!DayType = "N"
            End If
            
            If .optWeek.Value = True Then
                    rsSAVE!WeekType = "Y"
            Else: rsSAVE!WeekType = "N"
            End If

            bsaveRECORD = False

            rsSAVE.Update
            rsSAVE.Requery
  End With

Exit Sub

err:
    UpdateErrorMessage
    rsSAVE.CancelUpdate
    rsSAVE.Requery
End Sub


Private Sub cmdUpdate_Click()
        ValidateRELN
        
        If bsaveRECORD = True Then
            saveRecord
            enableButtons
            disableALLRECORD
            showDURATIONMODE
        End If


Exit Sub
End Sub

Private Sub cmdSearch_Click()
On Error GoTo err

        ListView1.Enabled = True
        ListView1.SetFocus

        Exit Sub

err:
            ErrorMessage

End Sub

Private Sub Form_Activate()
    LoadReinList
    disableALLRECORD
    DisplayRecord
End Sub




Sub LoadReinList()
On Error GoTo err
    
        With Screen.ActiveForm.ListView1
        
                .ListItems.Clear
                .ColumnHeaders.Clear
                .ColumnHeaders.Add , , "DurationMode", .Width / 4
                .ColumnHeaders.Add , , "Duration", .Width / 4
                .ColumnHeaders.Add , , "Prepared By", .Width / 4
                .ColumnHeaders.Add , , "Date Prepared", .Width / 4
                

                .View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "select * from ODASPDuration"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListItems.Add(, , CStr(rsLIST!DurationMode))
                        
                        If Not IsNull(rsLIST!DurationDescription) Then
                            MyList.SubItems(1) = CStr(rsLIST!DurationDescription)
                        End If
                        
                        If Not IsNull(rsLIST!Preparedby) Then
                            MyList.SubItems(2) = CStr(rsLIST!Preparedby)
                        End If

                        If Not IsNull(rsLIST!dateprepared) Then
                            MyList.SubItems(3) = CStr(rsLIST!dateprepared)
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


Sub ListViewOp()
    Me.txtDurationMode.Text = ListView1.SelectedItem
    SelectedListItem = ListView1.SelectedItem
    cmdEdit.Caption = "&Edit"
    'Me.txtLoanDescription.Text = ListView1.SelectedItem.SubItems(1)
    Call DisplayRecord

End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Call ListViewOp
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call ListViewOp
End Sub


Sub DisplayRecord()

On Error GoTo err

        'Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
        
        rsFind.Open "SELECT * FROM ODASPDuration WHERE DurationMode = '" & frmODASPDuration.txtDurationMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                    With frmODASPDuration
                            .txtDurationMode = rsFind!DurationMode
                            .txtDurationDescription = rsFind!DurationDescription
                            
                            If rsFind!DayType = "Y" Then
                                    .optDay.Value = True
                            Else: .optDay.Value = False
                            End If
                            
                            If rsFind!WeekType = "Y" Then
                                    .optWeek.Value = True
                            Else: .optWeek.Value = False
                            End If
                            
                            If rsFind!Status = "A" Then
                                    .chkstatus.Value = 1
                            Else: .chkstatus.Value = 0
                            End If
                            
                            
                    End With
                    
                    Edit = True
                End If

            End With

        Exit Sub

err:
            ErrorMessage
End Sub
