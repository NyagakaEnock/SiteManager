VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmALISPReinstatement 
   Caption         =   "Reinstatement Rates"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10080
   Icon            =   "frmALISPReinstatement.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmALISPReinstatement.frx":0442
   ScaleHeight     =   5430
   ScaleWidth      =   10080
   Begin VB.Frame Frame12 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
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
         Height          =   3735
         Left            =   7560
         TabIndex        =   8
         Top             =   1560
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
            Height          =   495
            Left            =   120
            Picture         =   "frmALISPReinstatement.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   15
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
            Height          =   495
            Left            =   120
            Picture         =   "frmALISPReinstatement.frx":0886
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
            Picture         =   "frmALISPReinstatement.frx":0988
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
            Picture         =   "frmALISPReinstatement.frx":0A8A
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
            Picture         =   "frmALISPReinstatement.frx":0B8C
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
            Picture         =   "frmALISPReinstatement.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3120
            Width           =   855
         End
      End
      Begin VB.Frame Frame13 
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   7335
         Begin VB.TextBox txtInterest 
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
            Height          =   400
            Left            =   4440
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtMonths 
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
            Height          =   400
            Left            =   960
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Interest Due"
            Height          =   255
            Left            =   2760
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblRelationshipCode 
            Caption         =   "Months"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   315
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
            BackColor       =   16777152
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   0
         End
      End
   End
End
Attribute VB_Name = "frmALISPReinstatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRELN As ADODB.Recordset, strRELN As String, bsaveRECORD As Boolean
Sub loadGRID()
        'Set RelationGrid.DataSource = rsRELN
End Sub

Sub clearRELN()
        txtMonths.Text = ""
        txtInterest.Text = ""
End Sub

Sub enableRELN()
        txtMonths.Locked = False
        txtInterest.Locked = False
End Sub

Sub disableRELN()
        txtMonths.Locked = True
        txtInterest.Locked = True
End Sub

Sub showRELN()
    With rsRELN
        txtMonths = !Months
        txtInterest = !Interest
        
    End With
End Sub


Private Sub cmdAdd_Click()
        
        ListView1.Enabled = False
        enableRELN
        clearRELN
        disableButtons
End Sub

Private Sub cmdCancel_Click()
        
        ListView1.Enabled = False
        enableButtons
        clearRELN
        disableRELN
End Sub


Private Sub cmdDelete_Click()
On Error GoTo err

If txtMonths.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            With rsRELN
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearRELN
                
                                         
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
                        enableRELN

                        'strQRE = InputBox("Enter The Months to search.", "Search Value")
    
                        rsFind.Open "SELECT * FROM ALISPReinstatement WHERE Months LIKE '" & SelectedListItem & "';", cnALIS, adOpenKeyset, adLockOptimistic

                        With rsFind
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                Else
                                        txtMonths = !Months
                                        txtInterest = !Interest
                                        Edit = True
                                End If
                        End With
        
                        If Edit Then
                                cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ALISPReinstatement WHERE Months = '" & txtMonths.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                            !Months = txtMonths
                            !Interest = txtInterest
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


Private Sub cmdFirstCode_Click(Index As Integer)

On Error GoTo err

        cmdUpdate.Enabled = False

        With rsRELN
        If .EOF And .BOF Then Exit Sub
    
                    Select Case Index
                                Case 0
                                    .MoveFirst
                                Case 1
                                    .MovePrevious
                                    If .BOF Then .MoveFirst
                                Case 2
                                    .MoveNext
                                    If .EOF Then .MoveLast
                                Case 3
                                    .MoveLast
                    End Select
        End With

                    showRELN
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ValidateRELN()
On Error GoTo err

        If txtMonths.Text = "" Then
                MsgBox "The Relationship Code  is Required"
                txtMonths.SetFocus
        ElseIf txtMonths.Text <= "" Then
                MsgBox "The Relationship is  required"
                txtMonths.SetFocus
        Else
                bsaveRECORD = True
        End If
Exit Sub

err:
ErrorMessage
End Sub

Private Sub SaveRECORD()
On Error GoTo err
      
   With rsRELN
        .AddNew
        !Months = txtMonths
        !Interest = txtInterest
        !Preparedby = CurrentUserName
        !dateprepared = Date
         .Update
         .Requery
  End With
Exit Sub

err:
    UpdateErrorMessage
    rsRELN.CancelUpdate
    rsRELN.Requery
End Sub


Private Sub cmdPrint_Click()
    Load frmParamRelationships
    frmParamRelationships.Show 1, ALISSysManager
End Sub

Private Sub cmdUpdate_Click()
        ValidateRELN
        
        If bsaveRECORD = True Then
            bsaveRECORD = False
            SaveRECORD
        End If
        
        enableButtons
        disableRELN
        cmdAdd.SetFocus


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
Call LoadReinList
End Sub

Private Sub Form_Load()

    Call OpenConnection
      
    Set rsRELN = New Recordset
            strRELN = "SELECT * from ALISPReinstatement;"

    rsRELN.Open strRELN, cnALIS, adOpenKeyset, adLockOptimistic

    disableRELN
    
   Call loadGRID

    Exit Sub
    
End Sub



Sub LoadReinList()
On Error GoTo err
    
        With Screen.ActiveForm.ListView1
        
                .ListItems.Clear
                .ColumnHeaders.Clear
                .ColumnHeaders.Add , , "Months", .Width / 7
                .ColumnHeaders.Add , , "interest", .Width / 7
                .ColumnHeaders.Add , , "Prepared By", .Width / 7
                .ColumnHeaders.Add , , "Date Prepared", .Width / 7
                

                .View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "select * from ALISPReinstatement"
                rsLIST.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListItems.Add(, , CStr(rsLIST!Months))
                        
                        If Not IsNull(rsLIST!Interest) Then
                            MyList.SubItems(1) = CStr(rsLIST!Interest)
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
    Me.txtMonths.Text = ListView1.SelectedItem
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
        'strQRE = InputBox("Enter The Month to search.", "Search Value")
        
        rsFind.Open "SELECT * FROM ALISPReinstatement WHERE Months = '" & SelectedListItem & "';", cnALIS, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                    txtMonths = !Months
                    txtInterest = !Interest
                    Edit = True
                End If

            End With

        Exit Sub

err:
            ErrorMessage
End Sub
