VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPDesignation 
   Caption         =   "Designation Management"
   ClientHeight    =   5325
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   10815
         Begin VB.OptionButton optHelper 
            Caption         =   "Helper"
            Height          =   255
            Left            =   9360
            TabIndex        =   24
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optCasual 
            Caption         =   "Casual"
            Height          =   255
            Left            =   7920
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optFlexiTeam 
            Caption         =   "FlexiTeam"
            Height          =   255
            Left            =   6480
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton OptsignWriter 
            Caption         =   "Signwriter"
            Height          =   255
            Left            =   4800
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optElectrician 
            Caption         =   "Electrician"
            Height          =   255
            Left            =   3360
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optDriver 
            Caption         =   "Driver"
            Height          =   255
            Left            =   2400
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optWelder 
            Caption         =   "Welder"
            Height          =   255
            Left            =   1320
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optSupervisor 
            Caption         =   "Supervisor"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   9615
         Begin MSComctlLib.ListView ListView1 
            Height          =   2895
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   5106
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
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   10815
         Begin VB.TextBox txtDesignationDescription 
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
            Height          =   315
            Left            =   4440
            TabIndex        =   11
            Top             =   240
            Width           =   4815
         End
         Begin VB.TextBox txtDesignationCode 
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
            Height          =   315
            Left            =   1320
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Left            =   3360
            TabIndex        =   13
            Top             =   300
            Width           =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Designation"
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   300
            Width           =   840
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   9840
         TabIndex        =   1
         Top             =   1680
         Width           =   1095
         Begin VB.CommandButton cmdSave 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPDesignation.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   615
            Width           =   855
         End
         Begin VB.CommandButton cmdAdd 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPDesignation.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edi&t"
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
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1365
            Width           =   855
         End
         Begin VB.CommandButton cmdRefresh 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPDesignation.frx":0204
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   990
            Width           =   855
         End
         Begin VB.CommandButton cmdSearch 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPDesignation.frx":0306
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1740
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPDesignation.frx":0408
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2115
            Width           =   855
         End
         Begin VB.CommandButton cmdPrintDesignations 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPDesignation.frx":050A
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   2490
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmODASPDesignation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_Click()
If rep = True Then Exit Sub
    If Screen.ActiveForm.txtDesignationCode.Text = "" Then
        MsgBox "There is no current record to delete", vbInformation, "Delete Information"
    ElseIf Screen.ActiveForm.txtDesignationDescription.Text = "" Then
        MsgBox "There is no current record to delete", vbInformation, "Delete Information"
    Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        Set rsDESIGNATION = New adodb.Recordset
        rsDESIGNATION.Open "SELECT * FROM ODASPDesignation WHERE DesignationCode='" & Me.txtDesignationCode.Text & "' ", cnPAY, adOpenKeyset, adLockOptimistic
        
        With rsDESIGNATION
            If .EOF And .BOF Then Exit Sub
            .Delete adAffectCurrent
            .Requery
             clearALLRECORD
             loadGRID
        End With
        
        End If
    End If
    Exit Sub
Myerr:
 ErrorMessage

End Sub

Private Sub cmdEdit_Click()
On Error GoTo err
    
    enableALLRECORD
    disableALLRECORD

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdPrintDesignations_Click()
    frmPWRDesignation.Show
End Sub

Private Sub cmdRefresh_Click()
    clearALLRECORD
    disableALLRECORD
    enableButtons
End Sub

Private Sub cmdsave_Click()
    rep = True
    validateDESIGNATION
    If bSaveRECORD = True Then
        bSaveRECORD = False
        saveDESIGNATION
        loadGRID
        enableButtons
        disableALLRECORD
    End If
End Sub
Private Sub cmdAdd_Click()
    rep = True
    enableALLRECORD
    clearALLRECORD
    disableButtons
End Sub


Sub validateDESIGNATION()
On Error GoTo err
        With frmODASPDesignation
                If .txtDesignationCode = "" Then
                MsgBox "Enter the Designation Code", vbInformation + vbOKOnly, "Validate Text"
                .txtDesignationCode.SetFocus
                
                ElseIf .txtDesignationDescription = "" Then
                MsgBox "Enter the Designation Description", vbInformation + vbOKOnly, "Validate Text"
                .txtDesignationDescription.SetFocus
                
                Else
                        bSaveRECORD = True
                End If
        End With
        
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadGRID()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Designation", .ListView1.Width / 10
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Supervisor", .ListView1.Width / 10
                .ListView1.ColumnHeaders.Add , , "Welder", .ListView1.Width / 10
                .ListView1.ColumnHeaders.Add , , "Driver", .ListView1.Width / 10
                .ListView1.ColumnHeaders.Add , , "Electrician", .ListView1.Width / 10
                .ListView1.ColumnHeaders.Add , , "SignWriter", .ListView1.Width / 10
                .ListView1.ColumnHeaders.Add , , "FlexiTeam", .ListView1.Width / 10
                .ListView1.ColumnHeaders.Add , , "Casual", .ListView1.Width / 10
                .ListView1.ColumnHeaders.Add , , "Helper", .ListView1.Width / 10

                .ListView1.View = lvwReport
                
                Dim rsList As adodb.Recordset
                Set rsList = New adodb.Recordset
                
                strSQL = "select DesignationCode, Description, PreparedBy, DatePrepared from ODASPDesignation"
                rsList.Open strSQL, cnPAY, adOpenKeyset, adLockOptimistic
                
                DF = rsList.RecordCount
                Dim MyList As ListItem
                           
                While Not rsList.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsList!DesignationCode))
                        
                        If Not IsNull(rsList!Description) Then
                            MyList.SubItems(1) = CStr(rsList!Description)
                        End If

                        If Not IsNull(rsList!Supervisor) Then
                                MyList.SubItems(2) = CStr(rsList!Supervisor)
                        End If
                        
                        If Not IsNull(rsList!Welder) Then
                                MyList.SubItems(3) = CStr(rsList!Welder)
                        End If
                        
                        If Not IsNull(rsList!Driver) Then
                                MyList.SubItems(4) = CStr(rsList!Driver)
                        End If
                        
                        If Not IsNull(rsList!Electrician) Then
                                MyList.SubItems(5) = CStr(rsList!Electrician)
                        End If
 
                        If Not IsNull(rsList!SignWriter) Then
                                MyList.SubItems(6) = CStr(rsList!SignWriter)
                        End If
                        
                        If Not IsNull(rsList!FlexiTeam) Then
                                MyList.SubItems(7) = CStr(rsList!FlexiTeam)
                        End If
                        
                        If Not IsNull(rsList!Casual) Then
                                MyList.SubItems(8) = CStr(rsList!Casual)
                        End If
                        
                        If Not IsNull(rsList!helper) Then
                                MyList.SubItems(9) = CStr(rsList!helper)
                        End If


                        rsList.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Private Sub saveDESIGNATION()
On Error GoTo err

    With frmODASPDesignation
        
        strSQL = "select * from ODASPDesignation where designationCode = '" & .txtDesignationCode.Text & "';"
        Set rsSAVE = New adodb.Recordset
        rsSAVE.Open strSQL, cnPAY, adOpenKeyset, adLockOptimistic

        If rsSAVE.EOF Or rsSAVE.BOF Then
                .AddNew
                rsSAVE!DesignationCode = Screen.ActiveForm.txtDesignationCode
                rsSAVE!DatePrepared = Date
                rsSAVE!Preparedby = CurrentUserName
        End If
        
        rsSAVE!Description = .txtDesignationDescription
        
        If .optCasual.Value = True Then
            rsSAVE!Casual = "Y"
        Else: rsSAVE!Casual = "N"
        End If
        
        If .optDriver.Value = True Then
            rsSAVE!Driver = "Y"
        Else: rsSAVE!Driver = "N"
        End If
        
        If .optElectrician.Value = True Then
                rsSAVE!Electrician = "Y"
        Else: rsSAVE!Electrician = "N"
        End If
        
        If .optFlexiTeam.Value = True Then
                rsSAVE!Electrician = "Y"
        Else: rsSAVE!Electrician = "N"
        End If
        
        If .optHelper.Value = True Then
                rsSAVE!helper = "Y"
        Else: rsSAVE!helper = "N"
        End If
        
        If .OptsignWriter.Value = True Then
                rsSAVE!SignWriter = "Y"
        Else: rsSAVE!SignWriter = "N"
        End If
        
        If .optSupervisor.Value = True Then
                rsSAVE!Supervisor = "Y"
        Else: rsSAVE!Supervisor = "N"
        End If
        
        If .optWelder.Value = True Then
                rsSAVE!Welder = "Y"
        Else: rsSAVE!Welder = "N"
        End If
        
        .Update
        .Requery
    
    
    End With
        
         rep = False

Exit Sub

err:
    If err.Number = -2147217842 Or err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
        rsDESIGNATION.CancelUpdate
        rsDESIGNATION.Requery
        cmdRefreshD.SetFocus
    Else
        ErrorMessage
    End If
End Sub



Private Sub cmdBrowseD_Click(Index As Integer)
On Error GoTo err

If rep = True Then Exit Sub
        With rsDESIGNATION
        If .EOF And .BOF Then Exit Sub
            Select Case Index
            Case 0
            .MoveFirst
            Case 1
            .MovePrevious
            If .BOF Then .MoveFirst
            'MsgBox "You are on the first record!"
            Case 2
            .MoveNext
            If .EOF Then .MoveLast
            'MsgBox "You are on the last record!"
            Case 3
            .MoveLast
            End Select
        End With
loadDESIGNATION

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadDESIGNATION()
On Error GoTo err
With rsDESIGNATION
        Screen.ActiveForm.txtDesignationCode.Text = !DesignationCode
        Screen.ActiveForm.txtDesignationDescription.Text = !Description
        Screen.ActiveForm.cboCompanyCode.Text = !CompanyCode & ""
End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub loadRECORD()
On Error GoTo err
    With frmODASPDesignation
    
        strSQL = "select * from ODASPDesignation where designationCode = '" & .txtDesignationCode.Text & "';"
        Set rsCONTROL = New adodb.Recordset
        rsCONTROL.Open strSQL, cnPAY, adOpenKeyset, adLockOptimistic
    
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
        .txtDesignationCode.Text = rsCONTROL!DesignationCode
        .txtDesignationDescription.Text = rsCONTROL!Description
            
        If !Supervisor = "Y" Then
                .optSupervisor.Value = True
        Else: .optSupervisor.Value = False
        End If
        
        If !Casual = "Y" Then
                .optCasual.Value = True
        Else: .optCasual.Value = False
        End If
  
        If !Driver = "Y" Then
                .optDriver.Value = True
        Else: .optDriver.Value = False
        End If
        
        If !Electrician = "Y" Then
                .optElectrician.Value = True
        Else: .optElectrician.Value = False
        End If
        
        If !FlexiTeam = "Y" Then
                .optFlexiTeam.Value = True
        Else: .optFlexiTeam.Value = False
        End If
        
        If !helper = "Y" Then
                .optHelper.Value = True
        Else: .optHelper.Value = False
        End If
        
        If !SignWriter = "Y" Then
                .OptsignWriter.Value = True
        Else: .OptsignWriter.Value = False
        End If
        
        If !Welder = "Y" Then
                .optWelder.Value = True
        Else: .optWelder.Value = False
        End If
 
 
    End With


rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdSearch_Click()
On Error GoTo ErrPCs
If rep = True Then Exit Sub
Dim strQRE As Variant

Dim rsFind As adodb.Recordset, Search As Boolean
Set rsFind = New adodb.Recordset

strQRE = InputBox("Enter Designation Code to search.", "Search Value")
If Len(strQRE) = 0 Then Exit Sub
   
rsFind.Open "SELECT * FROM ODASPDesignation WHERE DesignationCode = '" & strQRE & "';", cnPAY, adOpenKeyset, adLockOptimistic

With rsFind
    If .EOF And .BOF Then
        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
 
    Else
        Screen.ActiveForm.txtDesignationCode = !DesignationCode
        Screen.ActiveForm.txtDesignationDescription = !Description
        Screen.ActiveForm.cboCompanyCode = !CompanyCode & ""
        Search = True
    End If
End With

Exit Sub
ErrPCs:
    If err.Number = 40009 Then
        MsgBox "Record requested does not exist in the Database! Check your Entries.", vbInformation, "Searching."
        rsFind.Requery
        If rsFind.BOF Then Exit Sub
        rsFind.MoveFirst
    ElseIf err.Number = 3021 Then
        MsgBox "Requested record not found! Refresh the database and try the search again...or Check your entries.", vbInformation, "Searching."
        rsFind.Requery
        Screen.ActiveForm.txtDesignationCode.SetFocus
        If rsFind.BOF Then Exit Sub
        rsFind.MoveFirst
    Else
        MsgBox "ErrorMessage"
    End If

End Sub


Private Sub Form_Activate()
    disableALLRECORD
    
    loadRECORD
    loadGRID
    enableButtons
End Sub

Private Sub Form_Load()
    OpenConnection
End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err

        With frmODASPDesignation
            .txtDesignationCode.Text = Item.Text
            enableButtons
            loadRECORD
        End With

Exit Sub

err:
    ErrorMessage
End Sub

