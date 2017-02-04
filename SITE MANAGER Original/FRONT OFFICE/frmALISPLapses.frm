VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmALISPLapses 
   Caption         =   "Lapse Setup"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame6 
         Height          =   2775
         Left            =   5760
         TabIndex        =   16
         Top             =   120
         Width           =   4215
         Begin VB.TextBox txtLapseNotice 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   4
            Top             =   2160
            Width           =   2175
         End
         Begin VB.TextBox txtSecondNotice 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   1440
            TabIndex        =   3
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox txtPaymentMethod 
            BackColor       =   &H00FFFFC0&
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txtFirstNotice 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   1440
            TabIndex        =   2
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox txtPaymentMode 
            BackColor       =   &H00FFFFC0&
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label5 
            Caption         =   "Lapse Notice"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2220
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Second Notice"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1740
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Payment Method"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "First Notice"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1260
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Payment mode"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Payment Method"
         Height          =   2775
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   5535
         Begin MSComctlLib.ListView ListView2 
            Height          =   2415
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   4260
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
         Left            =   8880
         TabIndex        =   8
         Top             =   2880
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
            Picture         =   "frmALISPLapses.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
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
            Picture         =   "frmALISPLapses.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   12
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
            Picture         =   "frmALISPLapses.frx":0204
            Style           =   1  'Graphical
            TabIndex        =   11
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
            Picture         =   "frmALISPLapses.frx":0306
            Style           =   1  'Graphical
            TabIndex        =   10
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
            TabIndex        =   9
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
            Picture         =   "frmALISPLapses.frx":0408
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   600
            Width           =   855
         End
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
            Picture         =   "frmALISPLapses.frx":050A
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   8655
         Begin MSComctlLib.ListView ListView1 
            Height          =   2535
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
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
   End
End
Attribute VB_Name = "frmALISPLapses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    baddRECORD = True
    clearRECORD
    enableALLRECORD
    disableButtons
End Sub


Private Sub cmdCancel_Click()
    clearRECORD
    enableButtons
    disableALLRECORD
    baddRECORD = False
    beditRECORD = False
    bsearchRECORD = False
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Myerr

        If Screen.ActiveForm.PaymentMode.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
        Else
            If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
                    Dim rsDELETE As ADODB.Recordset
                    Set rsDELETE = New ADODB.Recordset
                    
                    rsDELETE.Open "Select * from ALISPReinsuranceParticipants where PaymentMode = '" & Screen.ActiveForm.txtPaymentMode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
             
                    With rsDELETE
                        If .EOF And .BOF Then Exit Sub
                        .Delete
                        .Requery
                         clearALLRECORD
                         loadGRID
                   End With
            End If
                '/* End if Msg Box
                
        End If
                '/* If txt = ""
rsDELETE.Close

Exit Sub

Myerr:
    ErrorMessage
End Sub

Private Sub cmdEdit_Click()
    editMYRECORD
End Sub

Private Sub cmdSearch_Click()
    searchMyRecord
End Sub

Private Sub cmdUpdate_Click()
    validateRECORD
    If bsaveRECORD = True Then
        SaveRECORD
        If bsaveRECORD = False Then
            enableButtons
            disableALLRECORD
            loadGRID
            loadPAYMENTMETHOD
        End If
    End If
    baddRECORD = False
    beditRECORD = False
End Sub

Private Sub validateRECORD()
On Error GoTo err

        bsaveRECORD = False
        
        With frmALISPLapses
            If .txtPaymentMode.Text <= "" Then
                    MsgBox "The Payment Mode is Invalid"
                    .txtPaymentMode.SetFocus

            ElseIf .txtLapseNotice.Text <= "" Then
                    MsgBox "The Lapse Notice is necessary"
                    .txtLapseNotice.SetFocus
            
            ElseIf .txtPaymentMethod.Text <= "" Then
                    MsgBox "The Payment Method is necessary"
                    .txtPaymentMethod.SetFocus
        
            ElseIf .txtFirstNotice.Text <= "" Then
                    MsgBox "The First Notice is necessary"
                    .txtFirstNotice.SetFocus
                    
            ElseIf .txtSecondNotice.Text <= "" Then
                    MsgBox "The Second Notice is Required"
                    .txtSecondNotice.SetFocus
            Else
                    bsaveRECORD = True
            End If
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveRECORD()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    strSQL = "Select * from ALISPLapses Where ALISPLapses.PaymentMode = '" & frmALISPLapses.txtPaymentMode.Text & "' and ALISPLapses.PaymentMethod = '" & frmALISPLapses.txtPaymentMethod.Text & "'; "
    rsSAVE.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
    
    With frmALISPLapses

            If rsSAVE.EOF Or rsSAVE.BOF Then
                    rsSAVE.AddNew
                    rsSAVE!PaymentMode = .txtPaymentMode
                    rsSAVE!PaymentMethod = .txtPaymentMethod.Text
                    rsSAVE!Preparedby = CurrentUserName
                    rsSAVE!dateprepared = Date
            End If
            
            rsSAVE!SecondNotice = .txtSecondNotice
            rsSAVE!FirstNotice = .txtFirstNotice.Text
            rsSAVE!LapseNotice = .txtLapseNotice

            bsaveRECORD = False
            rsSAVE.Update
            rsSAVE.Requery
    End With


Exit Sub

err:
    ErrorMessage
End Sub
Private Sub loadRECORD()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    strSQL = "Select * from ALISPLapses Where ALISPLapses.PaymentMode = '" & frmALISPLapses.txtPaymentMode.Text & "' and ALISPLapses.PaymentMethod = '" & frmALISPLapses.txtPaymentMethod.Text & "'; "
    rsCONTROL.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
    
    With frmALISPLapses
            If rsCONTROL.EOF Or rsCONTROL.BOF Then
                    MsgBox "The Record Entered does bot Exist", vbOKOnly
            Else

                    .txtPaymentMode = rsCONTROL!PaymentMode
                    .txtFirstNotice = rsCONTROL!FirstNotice
                    .txtPaymentMethod.Text = rsCONTROL!PaymentMethod
                    .txtSecondNotice.Text = rsCONTROL!SecondNotice
                    .txtLapseNotice.Text = rsCONTROL!LapseNotice & ""
            End If
            
    End With


Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadPAYMENTMETHOD()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                .ListView2.ColumnHeaders.Add , , "Method", .ListView2.Width / 2
                .ListView2.ColumnHeaders.Add , , "Description", .ListView2.Width / 2

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = " Select * FROM ALISPPaymentMethod ;"
                rsLIST.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set rsCONTROL = New ADODB.Recordset
                        strCONTROL = " Select * FROM ALISPLapses WHERE PaymentMode = '" & frmALISPLapses.txtPaymentMode & "' and PaymentMethod =  '" & rsLIST!PaymentMethod & "';"
                        rsCONTROL.Open strCONTROL, cnALIS, adOpenKeyset, adLockOptimistic
                        
                        If rsCONTROL.EOF Or rsCONTROL.BOF = True Then
                                
                                Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!PaymentMethod))
                                
                                If Not IsNull(rsLIST!PaymentMethodDescription) Then
                                    MyList.SubItems(1) = CStr(rsLIST!PaymentMethodDescription)
                                End If
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

Public Sub loadPaymentMode()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "PaymentMode", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = " Select PaymentMode, Description FROM ALISPPaymentMode;"
                rsLIST.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
            
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PaymentMode))
                    
                    If Not IsNull(rsLIST!Description) Then
                        MyList.SubItems(1) = CStr(rsLIST!Description)
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
                .ListView1.ColumnHeaders.Add , , "Mode", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Method", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "1st Notice", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "2nd Notice", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Lapse Notice", .ListView1.Width / 6
                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "select * from ALISPLapses where PaymentMode  = '" & frmALISPLapses.txtPaymentMode & "' order by PaymentMode, PaymentMethod "
                rsLIST.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PaymentMode))
                        
                        If Not IsNull(rsLIST!PaymentMethod) Then
                                MyList.SubItems(1) = CStr(rsLIST!PaymentMethod)
                        End If
                        
                        If Not IsNull(rsLIST!FirstNotice) Then
                                MyList.SubItems(2) = CStr(rsLIST!FirstNotice)
                        End If
                        
                        If Not IsNull(rsLIST!SecondNotice) Then
                                MyList.SubItems(3) = CStr(rsLIST!SecondNotice)
                        End If

                        If Not IsNull(rsLIST!LapseNotice) Then
                                MyList.SubItems(4) = CStr(rsLIST!LapseNotice)
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
    loadPAYMENTMETHOD
End Sub
Private Sub clearRECORD()
On Error GoTo err
    With frmALISPLapses
        .txtPaymentMethod.Text = ""
        .txtSecondNotice.Text = ""
        .txtLapseNotice.Text = ""
        .txtFirstNotice.Text = ""
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub loadDEFAULT()
On Error GoTo err
    With frmALISPLapses
        .txtFirstNotice.Text = "0"
        'CurrentPeriod
        .txtPaymentMethod.Text = PaymentMethod
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
            

            With frmALISPLapses
                    If .ListView1.Checkboxes = False Then
                            Exit Sub
                    Else
                        With frmALISPLapses
                                .txtPaymentMethod.Text = Item.SubItems(1)
                        End With
                    End If

                    frmALISPLapses.txtPaymentMode = Item.Text
                    disableALLRECORD
                    enableButtons
                    loadRECORD
            End With
        
        
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
        
        If bsearchRECORD <> True And beditRECORD <> True And baddRECORD <> True Then
                Item.Checked = False
                Exit Sub
        End If
        Dim i, j As Double
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView2.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                    If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                                Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                    End If
            Next i
            

            With frmALISPLapses
                    If .ListView2.Checkboxes = False Then
                            Exit Sub
                    Else
                        With frmALISPLapses
                                .txtPaymentMethod.Text = Item.Text
                        End With
                    End If
                    loadGRID
            End With
        
        
        End If
  
Exit Sub

err:
    ErrorMessage

End Sub

