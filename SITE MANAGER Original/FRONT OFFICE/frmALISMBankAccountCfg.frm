VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISMBankAccountCfg 
   Caption         =   "Cheque Issue Schedule Letter (Received)"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   11880
   Icon            =   "frmALISMBankAccountCfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin TabDlg.SSTab SSTabBankCheque 
         Height          =   6615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   11668
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Bank Cheque"
         TabPicture(0)   =   "frmALISMBankAccountCfg.frx":0442
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmdFirstCode(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdFirstCode(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdFirstCode(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdFirstCode(0)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame6"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame4"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Frame2"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "View Letters"
         TabPicture(1)   =   "frmALISMBankAccountCfg.frx":045E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "BeneficiaryGrid"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "frmALISMBankAccountCfg.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   9075
            TabIndex        =   29
            Top             =   1560
            Width           =   2415
            Begin VB.CommandButton cmdAddNew 
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
               Left            =   240
               TabIndex        =   36
               Top             =   240
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
               Left            =   240
               TabIndex        =   35
               Top             =   735
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
               Left            =   240
               TabIndex        =   34
               Top             =   1230
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
               Left            =   240
               TabIndex        =   33
               Top             =   1725
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
               Left            =   240
               TabIndex        =   32
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
               Left            =   240
               TabIndex        =   31
               Top             =   2715
               Width           =   2055
            End
            Begin VB.CommandButton cmdPrintletter 
               Caption         =   "&Pick Cheques"
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
               Left            =   240
               TabIndex        =   30
               Top             =   3210
               Width           =   2055
            End
         End
         Begin VB.Frame Frame4 
            Height          =   1215
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   11415
            Begin VB.TextBox txtcontactNames 
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
               Left            =   4920
               TabIndex        =   23
               Top             =   255
               Width           =   3255
            End
            Begin VB.ComboBox cboBankNo 
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
               TabIndex        =   22
               Top             =   720
               Width           =   2055
            End
            Begin VB.TextBox txttelephoneNo 
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
               Left            =   9480
               TabIndex        =   21
               Top             =   255
               Width           =   1575
            End
            Begin VB.CommandButton Command1 
               Caption         =   "&Print Schedule"
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
               Left            =   3960
               TabIndex        =   20
               Top             =   1920
               Width           =   2655
            End
            Begin VB.ComboBox txtAccountNo 
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
               Left            =   4920
               TabIndex        =   19
               Top             =   720
               Width           =   3255
            End
            Begin VB.TextBox txtReferenceNo 
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
               TabIndex        =   18
               Top             =   255
               Width           =   2055
            End
            Begin VB.Label lbNames 
               Caption         =   "Contact Names"
               Height          =   255
               Left            =   3600
               TabIndex        =   28
               Top             =   315
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Bank No"
               Height          =   195
               Left            =   240
               TabIndex        =   27
               Top             =   810
               Width           =   630
            End
            Begin VB.Label Label2 
               Caption         =   "Telephone No"
               Height          =   255
               Left            =   8400
               TabIndex        =   26
               Top             =   315
               Width           =   1215
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Account No"
               Height          =   255
               Left            =   3600
               TabIndex        =   25
               Top             =   780
               Width           =   855
            End
            Begin VB.Label lbPolicyNo 
               AutoSize        =   -1  'True
               Caption         =   "Reference"
               Height          =   195
               Left            =   240
               TabIndex        =   24
               Top             =   345
               Width           =   750
            End
            Begin VB.Line Line1 
               BorderWidth     =   2
               X1              =   120
               X2              =   10680
               Y1              =   1440
               Y2              =   1440
            End
         End
         Begin VB.Frame Frame6 
            Height          =   4095
            Left            =   120
            TabIndex        =   6
            Top             =   1560
            Width           =   8895
            Begin VB.TextBox txtsubject 
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1440
               TabIndex        =   11
               Top             =   360
               Width           =   7215
            End
            Begin VB.TextBox txttopnote 
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1440
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Top             =   1080
               Width           =   7215
            End
            Begin VB.TextBox txtbottomnote 
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   1440
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   9
               Top             =   2280
               Width           =   7215
            End
            Begin VB.TextBox txtDatePrepared 
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   6240
               TabIndex        =   8
               Top             =   3600
               Width           =   2415
            End
            Begin VB.TextBox txtPreparedBy 
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1440
               TabIndex        =   7
               Top             =   3600
               Width           =   2775
            End
            Begin VB.Label Label9 
               Caption         =   "Subject"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   420
               Width           =   975
            End
            Begin VB.Label Label10 
               Caption         =   "Top Note"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Label11 
               Caption         =   "Bottom Note"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   2520
               Width           =   1095
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Date Prepared"
               Height          =   195
               Left            =   5040
               TabIndex        =   13
               Top             =   3705
               Width           =   1035
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Prepared By"
               Height          =   195
               Left            =   240
               TabIndex        =   12
               Top             =   3690
               Width           =   870
            End
         End
         Begin VB.CommandButton cmdFirstCode 
            Height          =   375
            Index           =   0
            Left            =   1395
            Picture         =   "frmALISMBankAccountCfg.frx":0496
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   5880
            Width           =   1695
         End
         Begin VB.CommandButton cmdFirstCode 
            Height          =   375
            Index           =   1
            Left            =   3090
            Picture         =   "frmALISMBankAccountCfg.frx":08D8
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   5880
            Width           =   1575
         End
         Begin VB.CommandButton cmdFirstCode 
            Height          =   375
            Index           =   2
            Left            =   4665
            Picture         =   "frmALISMBankAccountCfg.frx":0D1A
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   5880
            Width           =   1455
         End
         Begin VB.CommandButton cmdFirstCode 
            Height          =   375
            Index           =   3
            Left            =   6120
            Picture         =   "frmALISMBankAccountCfg.frx":115C
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   5880
            Width           =   1815
         End
         Begin MSDataGridLib.DataGrid BeneficiaryGrid 
            Height          =   6015
            Left            =   -74880
            TabIndex        =   37
            Top             =   480
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   10610
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
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
               Weight          =   700
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
   End
End
Attribute VB_Name = "frmALISMBankAccountCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCode As ADODB.Recordset, strcode, iReferenceNo As String, bAddNew As Boolean

Sub ClearControls()
On Error GoTo err
With frmALISMBankAccountCfg
    .cboBankNo.Text = ""
    .txtcontactNames.Text = ""
    .txtDatePrepared.Text = ""
    .txtPreparedBy.Text = ""
    .txtsubject.Text = ""
    .txttopnote.Text = ""
    .txtbottomnote.Text = ""
    .txtReferenceNo.Text = ""
    .txtTelephoneNo.Text = ""
    .txtAccountNo.Text = ""
   
End With
    Exit Sub

err:
ErrorMessage
End Sub

Sub EnableControls()
On Error GoTo err
With frmALISMBankAccountCfg
    .cboBankNo.Locked = False
    .txtcontactNames.Locked = False
    .txtDatePrepared.Locked = False
    .txtPreparedBy.Locked = False
    .txtsubject.Locked = False
    .txttopnote.Locked = False
    .txtbottomnote.Locked = False
    .txtAccountNo.Locked = False
    

End With
     Exit Sub

err:
ErrorMessage
End Sub

Sub DisableControls()
On Error GoTo err

   With frmALISMBankAccountCfg
    .cboBankNo.Locked = True
    .txtcontactNames.Locked = True
    .txtAccountNo.Locked = True
    .txtDatePrepared.Locked = True
    .txtPreparedBy.Locked = True
    .txtsubject.Locked = True
    .txttopnote.Locked = True
    .txtbottomnote.Locked = True
End With
Exit Sub

err:
ErrorMessage
End Sub

Sub ShowCode()

On Error GoTo err
    With RsCode
    cboBankNo.Text = !BankNo & ""
    txtDatePrepared = !dateprepared & ""
    txtPreparedBy = !Preparedby & ""
    txtAccountNo = !AccountNo & ""
    txtsubject = !Subject & ""
    txttopnote = !TopNote & ""
    txtbottomnote = !BottomNote & ""
End With
    Exit Sub

err:
ErrorMessage

End Sub

Private Sub DisableCommandButtons()
On Error GoTo err
    cmdUpdate.Enabled = True
    cmdAddNew.Enabled = False
    cmdSearch.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdPrintletter.Enabled = False
    Exit Sub

err:
ErrorMessage
End Sub
Private Sub EnableCommandButtons()
On Error GoTo err
    cmdUpdate.Enabled = False
    cmdAddNew.Enabled = True
    cmdSearch.Enabled = True
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdCancel.Enabled = True
    Exit Sub

err:
ErrorMessage
End Sub


Private Sub cbobankNo_GotFocus()
On Error GoTo err
cboBankNo.Clear
Dim i As Integer
Dim rsloadbank As ADODB.Recordset
Set rsloadbank = New ADODB.Recordset
rsloadbank.Open "Select * from ALISPBankAccount", cnALIS, adOpenKeyset, adLockOptimistic

For i = 1 To rsloadbank.RecordCount
    cboBankNo.AddItem rsloadbank!BankNo
    rsloadbank.MoveNext
Next i
Set rsloadbank = Nothing
Exit Sub
err:
ErrorMessage
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo Myerr
        bAddNew = True
        ClearControls
        EnableControls
        DisableCommandButtons
        cboBankNo.FontBold = True
        txtDatePrepared.Text = Date
        Exit Sub

Myerr:
        ErrorMessage
End Sub


Private Sub cmdCancel_Click()
On Error GoTo Myerr
        EnableCommandButtons
        ClearControls
        DisableControls
         Exit Sub

Myerr:
        ErrorMessage
End Sub

Private Sub cmdDelete_Click()

On Error GoTo Myerr

If txtReferenceNo.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
ElseIf txtcontactNames.Text = "" Then
            MsgBox "There is no current record", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            With RsCode
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                ClearControls
                
                                
            End With
    End If
        
End If
        
Exit Sub

Myerr:
    ErrorMessage

End Sub

Private Sub cmdedit_Click()

On Error GoTo Myerr

Dim strQRE As Variant
Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset

        Select Case cmdEdit.Caption
                Case "&Edit"
                        EnableControls

                        strQRE = InputBox("Enter Reference No  to search.", "Search Value")
    
                        rsFind.Open "SELECT * FROM ALISMBankAccountCfg WHERE Reference = '" & strQRE & "'and ReceivedPayed='" & "Received" & "';", cnALIS, adOpenKeyset, adLockOptimistic

                        With rsFind
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                Else
                                    cboBankNo.Text = !BankNo
                                    txtReferenceNo.Text = !reference
                                    txtDatePrepared = !dateprepared
                                    txtPreparedBy = !Preparedby
                                    txtsubject = !Subject
                                    txttopnote = !TopNote
                                    txtbottomnote = !BottomNote
                                    txtAccountNo.Text = !AccountNo & ""
                                    Edit = True
                                End If
     
                        End With
                        
                        
                        If Edit Then
                                cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ALISMBankAccountCfg WHERE Reference = '" & txtReferenceNo.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                                     !BankNo = cboBankNo.Text
                                     !dateprepared = txtDatePrepared.Text
                                     !Preparedby = txtPreparedBy.Text
                                     !Subject = txtsubject.Text
                                     !TopNote = txttopnote.Text
                                     !BottomNote = txtbottomnote.Text
                                     !AccountNo = txtAccountNo.Text
                            .Update
                            .Requery
                            Edit = False
                    End With
                
                    ClearControls
                    cmdEdit.Caption = "&Edit"
            Case Else
        
            Exit Sub

        End Select

Exit Sub

Myerr:
            UpdateErrorMessage
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
    rsFinder.CancelUpdate
    rsFinder.Requery
End If
End Sub
Sub LoadDefaults()

On Error GoTo err

    Dim rsDEFA As ADODB.Recordset, strDEFA, LSetupCode As String
    Set rsDEFA = New Recordset
    
    LSetupCode = 0
    
    strDEFA = "Select * from ALISMReference where referenceNo LIKE  '" & iReferenceNo & "';"
    rsDEFA.Open strDEFA, cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsDEFA
        If .EOF And .BOF Then Exit Sub
            txtcontactNames = Trim(!OtherNames) + " " + Trim(!surname)
            txtDatePrepared = Date
   End With

rsDEFA.Close
strDEFA = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdFirstCode_Click(Index As Integer)
On Error GoTo Myerr

cmdUpdate.Enabled = False
With RsCode
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

ShowCode
DisableControls
Call LoadGrid
Exit Sub
Myerr:
    ErrorMessage
End Sub

Private Sub ValidateData()

On Error GoTo err
        If cboBankNo.Text = "" Then
                MsgBox "BankNo is Required"
                cboBankNo.SetFocus
      
                            
        ElseIf txtcontactNames.Text = "" Then
                MsgBox "The Names is  required"
                txtcontactNames.SetFocus
        
      
        ElseIf txtTelephoneNo.Text = "" Then
                MsgBox "The Telephone No is  required"
                txtTelephoneNo.SetFocus
           
        ElseIf txtAccountNo.Text = "" Then
                MsgBox "The Telephone No is  required"
                txtAccountNo.SetFocus
            
        Else
                GoTo save
        End If
save:
            With RsCode
                    .AddNew
                                     
                                     !BankNo = cboBankNo.Text
                                     !dateprepared = Date
                                     !Preparedby = CurrentUserName
                                     !AccountNo = txtAccountNo.Text
                                     !ReceivedPayed = "Received"
                                     If txtsubject.Text = "" Then
                                        !Subject = "RE: CHEQUE RECEIPT SCHEDULE A/C NO"
                                     Else
                                        !Subject = txtsubject.Text
                                     End If
                                     If txttopnote = "" Then
                                        !TopNote = " Please we have Received the following cheques as Payments:"
                                     Else
                                        !TopNote = txttopnote.Text
                                     End If
                                     If txtbottomnote.Text = "" Then
                                         !BottomNote = "Also note that the Cheque Receipt list can be signed by ONE signatory"
                                     Else
                                        !BottomNote = txtbottomnote.Text
                                     End If
                    .Update
            End With
     If frmALISMBankAccountCfg.txtReferenceNo.Text = "" Then
    MsgBox " The reference is required before proceeding", vbCritical
    Exit Sub
Else
Load frmALISMPickCheques
frmALISMPickCheques.Show 1, Me
End If
            
        Exit Sub
err:
            UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            RsCode.CancelUpdate
            RsCode.Requery
        End If
  
End Sub


Private Sub cmdPrintletter_Click()

On Error GoTo err
If frmALISMBankAccountCfg.txtReferenceNo.Text = "" Then
    MsgBox " The reference is required before proceeding", vbCritical
    Exit Sub
Else
Load frmALISMPickCheques
frmALISMPickCheques.Show 1, Me
End If
Exit Sub
err:
ErrorMessage
End Sub

Private Sub cmdUpdate_Click()

        generateReference
        ValidateData
        Call LoadGrid
        EnableCommandButtons
        DisableControls
End Sub

Private Sub generateReference()
On Error GoTo err
        Dim rsLAST As ADODB.Recordset, strLAST As String
        
        Set rsLAST = New Recordset
      
        strLAST = "SELECT * FROM ALISPLastNumber;"
        rsLAST.Open strLAST, cnALIS, adOpenKeyset, adLockOptimistic

                 With rsLAST
                
                frmALISMBankAccountCfg.txtReferenceNo = !reference
                !reference = !reference + 1
                    .Update
            End With

Exit Sub

err:
    UpdateErrorMessage
End Sub
Private Sub cmdSearch_Click()

On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
        
        strQRE = InputBox("Enter ReferenceNo to search.", "Search Value")
        rsFind.Open "SELECT * FROM ALISMBankAccountCfg  WHERE Reference = '" & strQRE & "'and ReceivedPayed='" & "Received" & "';", cnALIS, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                            ClearControls
                                    cboBankNo.Text = !BankNo
                                    txtReferenceNo.Text = !reference
                                    txtDatePrepared.Text = !dateprepared
                                    txtPreparedBy.Text = !Preparedby
                                    txtsubject.Text = !Subject & ""
                                    txttopnote.Text = !TopNote & ""
                                    txtbottomnote.Text = !BottomNote & ""
                                    txtAccountNo.Text = !AccountNo & ""
                                    Me.cmdPrintletter.Enabled = True
                                    Call cbobankNo_LostFocus
                            Edit = True
                End If
                'Call LoadClaimDetails
                            

        End With
        
        Exit Sub

Myerr:
            ErrorMessage

End Sub



Private Sub Form_Load()

On Error GoTo Myerr

    Call OpenConnection
      
            Set RsCode = New Recordset
            strcode = "SELECT * from ALISMBankAccountCfg"

   RsCode.Open strcode, cnALIS, adOpenKeyset, adLockOptimistic
   
    
    DisableControls
    
    Call LoadGrid

    cmdUpdate.Enabled = False

On Error Resume Next
Me.Top = 800
Me.Left = 600
    Exit Sub
    
Myerr:
    ErrorMessage

End Sub

Private Sub cboCauseCode_KeyPress(KeyAscii As Integer)

On Error GoTo err
          KeyAscii = 0
          Exit Sub

err:
ErrorMessage
End Sub

Private Sub LoadGrid()
    
On Error GoTo err

    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New Recordset

   rsGRID.Open "SELECT * FROM ALISMBankAccountCfg where ReceivedPayed='" & "Received" & "';", cnALIS, adOpenKeyset, adLockOptimistic
   Set BeneficiaryGrid.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cbobankNo_LostFocus()
On Error GoTo err
    Dim rscontact As ADODB.Recordset
    Set rscontact = New ADODB.Recordset
    rscontact.Open "Select * from ALISPBank where BankNo= '" & frmALISMBankAccountCfg.cboBankNo & "';", cnALIS, adOpenKeyset, adLockOptimistic
    If rscontact.EOF Or rscontact.BOF Then Exit Sub
    txtcontactNames.Text = rscontact!Contactname
    txtTelephoneNo.Text = rscontact!contactTelephoneNo
    Set rscontact = Nothing
    Exit Sub
err:
    ErrorMessage

End Sub



Private Sub txtAccountNo_GotFocus()
On Error GoTo err

txtAccountNo.Clear
Dim i As Integer
Dim rsloadbankacc As ADODB.Recordset
Set rsloadbankacc = New ADODB.Recordset
rsloadbankacc.Open "Select AccountNo from ALISPBankAccount where BankNo='" & frmALISMBankAccountCfg.cboBankNo.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
For i = 1 To rsloadbankacc.RecordCount
    txtAccountNo.AddItem rsloadbankacc!AccountNo
    rsloadbankacc.MoveNext
Next i
Set rsloadbankacc = Nothing

Exit Sub
err:
ErrorMessage

End Sub
