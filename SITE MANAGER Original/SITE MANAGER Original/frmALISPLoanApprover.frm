VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmALISPLoanApprover 
   Caption         =   "Loan Approvers"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   11340
   Icon            =   "frmALISPLoanApprover.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   11340
   Begin VB.Frame Frame12 
      Height          =   6615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   3135
         Left            =   120
         TabIndex        =   36
         Top             =   2640
         Width           =   4215
         Begin MSComctlLib.ListView ListView1 
            Height          =   2775
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   4895
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
            BackColor       =   16761024
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
         TabIndex        =   33
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
            TabIndex        =   34
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
            TabIndex        =   4
            Top             =   255
            Width           =   3375
         End
         Begin VB.Label Label5 
            Caption         =   "Operation "
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   330
            Width           =   1215
         End
      End
      Begin VB.Frame fraCButtons 
         Height          =   3855
         Index           =   6
         Left            =   8880
         TabIndex        =   25
         Top             =   2640
         Width           =   2295
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   3240
            Width           =   2055
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   28
            Top             =   1725
            Width           =   2055
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Update"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   27
            Top             =   2220
            Width           =   2055
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   26
            Top             =   2715
            Width           =   2055
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Properties"
         Height          =   1815
         Left            =   120
         TabIndex        =   20
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
            TabIndex        =   7
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   3
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
            TabIndex        =   10
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
            TabIndex        =   6
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
            TabIndex        =   8
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
            TabIndex        =   9
            Top             =   1215
            Width           =   2535
         End
         Begin VB.Label Label6 
            Caption         =   "Re-Enter Password"
            Height          =   195
            Left            =   7320
            TabIndex        =   37
            Top             =   810
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Limit"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Date Retired"
            Height          =   195
            Left            =   4560
            TabIndex        =   31
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Left            =   4560
            TabIndex        =   30
            Top             =   810
            Width           =   855
         End
         Begin VB.Label lblBenefitCode 
            Caption         =   "User Code"
            Height          =   315
            Left            =   240
            TabIndex        =   24
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Name"
            Height          =   195
            Left            =   3960
            TabIndex        =   23
            Top             =   330
            Width           =   855
         End
         Begin VB.Label lblStatus 
            Caption         =   "Status"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   1290
            Width           =   735
         End
         Begin VB.Label lblDateAssigned 
            Caption         =   "Date Assigned"
            Height          =   195
            Left            =   7320
            TabIndex        =   21
            Top             =   1260
            Width           =   1215
         End
      End
      Begin VB.Frame Frame14 
         Height          =   3135
         Left            =   4440
         TabIndex        =   19
         Top             =   2640
         Width           =   4335
         Begin MSDataGridLib.DataGrid AuthorizerGrid 
            Height          =   2775
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   4895
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame frabrowse 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   5760
         Width           =   8655
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   3
            Left            =   5400
            Picture         =   "frmALISPLoanApprover.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   2
            Left            =   4080
            Picture         =   "frmALISPLoanApprover.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   1
            Left            =   2640
            Picture         =   "frmALISPLoanApprover.frx":0B8E
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   0
            Left            =   1200
            Picture         =   "frmALISPLoanApprover.frx":0FD0
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmALISPLoanApprover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsLOANAPPROVERS As cLoanApprovers
Public rsLOAN As ADODB.Recordset

Private Sub cboOperationType_GotFocus()
        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.operationtypeGOTFOCUS
        Set rsLOANAPPROVERS = Nothing

End Sub
Private Sub cboOperationType_KeyPress(Index As Integer)
On Error GoTo err

        Set rsLOANAPPROVERS = New cLoanApprovers
        Set rsLOANAPPROVERS = Nothing

Exit Sub

err:
    ErrorMessage

End Sub
Private Sub cboOperationType_LostFocus()
        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.operationtypeLOSTFOCUS
        Set rsLOANAPPROVERS = Nothing
        GetUserCode

End Sub
Private Sub cboStatus_gotFocus()
        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.statusGOTFOCUS
        Set rsLOANAPPROVERS = Nothing
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

        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.statusLOSTFOCUS
        Set rsLOANAPPROVERS = Nothing

Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cmdAdd_Click()
        ClearListview
        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.AddRECORD
        Set rsLOANAPPROVERS = Nothing
End Sub
'Private Sub UpdateALLRECORDS()
'
'        'ValidateData
'        If bsaveRECORD = True Then
'                updateRECORD
'        End If
'End Sub

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
On Error GoTo err
        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.cancelRECORD
        Set rsLOANAPPROVERS = Nothing
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdDelete_Click()
On Error GoTo err

        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.DeleteRecord
        Set rsLOANAPPROVERS = Nothing

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdEdit_Click()
On Error GoTo err

        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.EditRecord
        Set rsLOANAPPROVERS = Nothing

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdFirstCode_Click(Index As Integer)
On Error GoTo err

        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.browseRECORD (Index)
        Set rsLOANAPPROVERS = Nothing

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdSearch_Click()
On Error GoTo err
                
        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.SearchRECORD
        Set rsLOANAPPROVERS = Nothing

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub GetUserCode()
On Error GoTo err
    
        With frmALISPLoanApprover
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Staff", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "User Name", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "All Names", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open "SELECT AdminUserRegister.StaffIdNo,AdminUserRegister.UserName, AdminUserRegister.AllNames FROM AdminUserRegister ;", cnALIS, adOpenKeyset, adLockOptimistic
                
                DF = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!staffidno))
                            If Not IsNull(rsLIST!UserName) Then
                                MyList.SubItems(1) = CStr(rsLIST!UserName)
                            End If
                            
                            If Not IsNull(rsLIST!allnames) Then
                                    MyList.SubItems(2) = CStr(rsLIST!allnames)
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

Private Sub cmdUpdate_Click()
        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.updateRECORD
        Set rsLOANAPPROVERS = Nothing
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
        'Set rsLOANAPPROVERS = New cLoanApprovers
        'Call rsLOANAPPROVERS.disableRECORD
        
       
        'Set rsLOAN = New Recordset
        'strLOAN = "SELECT * from ALISPLoanApprover;"

        'rsLOAN.Open strLOAN, cnALIS, adOpenKeyset, adLockOptimistic

        'rsLOANAPPROVERS.LoadGrid

        
        
        
Exit Sub
err:

ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        
        Set rsLOANAPPROVERS = New cLoanApprovers
        rsLOANAPPROVERS.clearRECORD
        Set rsLOANAPPROVERS = Nothing
        

        Dim strRequirementcode, StrLOAD As String
        Dim rsLOAD As ADODB.Recordset
        Set rsLOAD = New ADODB.Recordset
                
       
        strRequirementcode = Item.Text
        frmALISPLoanApprover.txtUserCode.Text = Item.Text

        StrLOAD = "Select * from ALISPLoanApprover, AdminUserRegister where ALISPLoanApprover.OperationType = '" & frmALISPLoanApprover.cboOperationType & "' and ALISPLoanApprover.UserCode = AdminUserRegister.StaffIDNo ;"
        rsLOAD.Open StrLOAD, cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsLOAD
                If .BOF Or .EOF Then
                        Dim rsNEW As ADODB.Recordset, strNEW As String
                        Set rsNEW = New ADODB.Recordset
 
                        strNEW = "Select * from AdminUserRegister where StaffIDNO = '" & frmALISPLoanApprover.txtUserCode & "';"
                        rsNEW.Open strNEW, cnALIS, adOpenKeyset, adLockOptimistic
                        
                        With rsNEW
                                If .BOF Or .EOF Then Exit Sub
                                frmALISPLoanApprover.txtDateAssigned.Text = Date
                                frmALISPLoanApprover.txtLimitAmount.Text = 0
                                frmALISPLoanApprover.txtNames.Text = !UserName
                                frmALISPLoanApprover.txtUserCode.Text = !staffidno
                                frmALISPLoanApprover.cboStatus = "A"
                        End With
                        
                        rsNEW.Close
                        strNEW = ""
                Else
                        frmALISPLoanApprover.txtDateAssigned.Text = !DateAssigned
                        frmALISPLoanApprover.txtLimitAmount.Text = !LimitAmount & ""
                        frmALISPLoanApprover.txtNames.Text = !UserName
                        frmALISPLoanApprover.txtUserCode.Text = !staffidno & ""
                        frmALISPLoanApprover.cboStatus = !Status & ""
                End If
         End With
rsLOAD.Close
StrLOAD = ""
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
