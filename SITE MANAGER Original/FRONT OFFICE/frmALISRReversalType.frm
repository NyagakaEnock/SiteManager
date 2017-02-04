VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmALISPReversalType 
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8895
      Begin VB.Frame Frame3 
         Height          =   2895
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   7455
         Begin MSComctlLib.ListView ListView1 
            Height          =   2535
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   4471
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
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   8655
         Begin VB.TextBox txtReversalDescription 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   720
            Width           =   5775
         End
         Begin VB.TextBox txtReversalType 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Description"
            Height          =   255
            Left            =   4920
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Reversal Type"
            Height          =   255
            Left            =   960
            TabIndex        =   12
            Top             =   300
            Width           =   1215
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
         Height          =   2895
         Left            =   7680
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
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
            Picture         =   "frmALISRReversalType.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2520
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
            Picture         =   "frmALISRReversalType.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2115
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
            Picture         =   "frmALISRReversalType.frx":0204
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1740
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
            Picture         =   "frmALISRReversalType.frx":0306
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1365
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
            TabIndex        =   4
            Top             =   990
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
            Picture         =   "frmALISRReversalType.frx":0408
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   600
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
            Height          =   375
            Left            =   120
            Picture         =   "frmALISRReversalType.frx":050A
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmALISPReversalType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddNew_Click()
    baddRECORD = True
    clearALLRECORD
    enableALLRECORD
    disableButtons
End Sub

Private Sub cmdCancel_Click()
    enableButtons
    disableALLRECORD
    baddRECORD = False
    beditRECORD = False
    bsearchRECORD = False
End Sub

Private Sub cmdDelete_Click()
''''On Error GoTo Myerr

        If Screen.ActiveForm.ReversalType.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
        Else
            If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
                    Dim rsDELETE As ADODB.Recordset
                    Set rsDELETE = New ADODB.Recordset
                    
                    rsDELETE.Open "Select * from ALISPReinsuranceParticipants where ReversalType = '" & Screen.ActiveForm.txtReversalType.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
             
                    With rsDELETE
                        If .EOF And .BOF Then Exit Sub
                        .Delete
                        .Requery
                         clearALLRECORD
                         loadGRID
                   End With
            End If
                
        End If
rsDELETE.Close

Exit Sub

Myerr:
    ErrorMessage
End Sub

Private Sub cmdEdit_Click()
    EditMyRecord
End Sub

Private Sub cmdSearch_Click()
    searchMyRecord
End Sub

Private Sub cmdUpdate_Click()
    validateRECORD
    If bsaveRECORD = True Then
        saveRecord
        If bsaveRECORD = False Then
            enableButtons
            disableALLRECORD
            loadGRID
        End If
    End If
    baddRECORD = False
    beditRECORD = False
End Sub

Private Sub validateRECORD()
On Error GoTo err

        bsaveRECORD = False
        
        With frmALISPReversalType
            If .txtReversalType.Text <= "" Then
                    MsgBox "The Reversal Type is Invalid"
                    .txtReversalType.SetFocus
            
            ElseIf .txtReversalDescription.Text <= "" Then
                    MsgBox "The Payment Method is necessary"
                    .txtReversalDescription.SetFocus
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
    
    Set rsSAVE = New ADODB.Recordset
    strSQL = "Select * from ALISPReversalType Where ALISPReversalType.ReversalType = '" & frmALISPReversalType.txtReversalType.Text & "' and ALISPReversalType.ReversalDescription = '" & frmALISPReversalType.txtReversalDescription.Text & "'; "
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With frmALISPReversalType

            If rsSAVE.EOF Or rsSAVE.BOF Then
                    rsSAVE.AddNew
                    rsSAVE!ReversalType = .txtReversalType
                    rsSAVE!Preparedby = CurrentUserName
                    rsSAVE!dateprepared = Date
            End If
            
            rsSAVE!ReversalDescription = .txtReversalDescription.Text

            bsaveRECORD = False
            rsSAVE.Update
    End With


Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadRECORD()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    strSQL = "Select * from ALISPReversalType Where ALISPReversalType.ReversalType = '" & frmALISPReversalType.txtReversalType.Text & "' and ALISPReversalType.ReversalDescription = '" & frmALISPReversalType.txtReversalDescription.Text & "'; "
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With frmALISPReversalType
            If rsCONTROL.EOF Or rsCONTROL.BOF Then
                    MsgBox "The Record Entered does bot Exist", vbOKOnly
            Else

                    .txtReversalType = rsCONTROL!ReversalType
                    .txtReversalDescription.Text = rsCONTROL!ReversalDescription
            End If
            
    End With


Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadReversalType()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "ReversalType", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = " Select ReversalType, ReversalDescription FROM ALISPReversalType;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
            
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReversalType))
                    
                    If Not IsNull(rsLIST!ReversalDescription) Then
                        MyList.SubItems(1) = CStr(rsLIST!ReversalDescription)
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

Public Sub loadGRID()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Reversal Type", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2
                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "select * from ALISPReversalType order by ReversalType "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReversalType))
                        
                        If Not IsNull(rsLIST!ReversalDescription) Then
                                MyList.SubItems(1) = CStr(rsLIST!ReversalDescription)
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

Private Sub Form_Activate()
    enableButtons
    disableALLRECORD
    loadGRID
End Sub
Private Sub clearRECORD()
On Error GoTo err
    With frmALISPReversalType
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub LoadDEFAULT()
On Error GoTo err
    With frmALISPReversalType
        .txtReversalDescription.Text = ReversalDescription
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub Form_Load()
    bsearchRECORD = False
    beditRECORD = False
    baddRECORD = False
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
        
        If bsearchRECORD <> True And beditRECORD <> True Then
                Item.Checked = False
                Exit Sub
        End If
        Dim i, j As Double
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                    If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                                Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                    End If
            Next i

            With frmALISPReversalType
                    If .ListView1.Checkboxes = False Then
                            Exit Sub
                    Else
                        With frmALISPReversalType
                                .txtReversalDescription.Text = Item.SubItems(1)
                        End With
                    End If

                    frmALISPReversalType.txtReversalType = Item.Text
                    disableALLRECORD
                    enableButtons
                    loadRECORD
            End With
        
        
        End If
  
Exit Sub

err:
    ErrorMessage

End Sub

