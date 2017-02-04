VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISMBankAccount 
   Caption         =   "Bank Account"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10965
   Icon            =   "frmALISMBankAccount.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10965
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
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
         Height          =   495
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   4080
         Width           =   6615
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
         Height          =   495
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   3600
         Width           =   6615
      End
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
         Height          =   375
         Left            =   1320
         TabIndex        =   36
         Top             =   2760
         Width           =   6615
      End
      Begin VB.CommandButton cmdFirstCode 
         Height          =   375
         Index           =   3
         Left            =   6045
         Picture         =   "frmALISMBankAccount.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5040
         Width           =   1815
      End
      Begin VB.CommandButton cmdFirstCode 
         Height          =   375
         Index           =   2
         Left            =   4590
         Picture         =   "frmALISMBankAccount.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdFirstCode 
         Height          =   375
         Index           =   1
         Left            =   3015
         Picture         =   "frmALISMBankAccount.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton cmdFirstCode 
         Height          =   375
         Index           =   0
         Left            =   1320
         Picture         =   "frmALISMBankAccount.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   50
         TabIndex        =   23
         Top             =   1560
         Width           =   7935
         Begin VB.TextBox txtSignatorytittle 
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
            Left            =   2280
            TabIndex        =   26
            Top             =   615
            Width           =   5535
         End
         Begin VB.TextBox txtSignatoryNames 
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
            Height          =   375
            Left            =   2280
            TabIndex        =   24
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label Label8 
            Caption         =   "Signatory Tittle"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   615
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Signatory Names"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtDetails 
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
         Height          =   495
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   3120
         Width           =   6615
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
         Left            =   5400
         TabIndex        =   17
         Top             =   4590
         Width           =   2535
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
         Left            =   1320
         TabIndex        =   15
         Top             =   4590
         Width           =   2775
      End
      Begin VB.Frame Frame4 
         Height          =   1575
         Left            =   50
         TabIndex        =   11
         Top             =   0
         Width           =   10815
         Begin VB.TextBox txtAccountNo 
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
            Height          =   375
            Left            =   7800
            TabIndex        =   34
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtContactDesignation 
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
            Height          =   375
            Left            =   1800
            TabIndex        =   19
            Top             =   1080
            Width           =   3975
         End
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
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox txtBankNo 
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
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox cboChequeNo 
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
            Left            =   6720
            TabIndex        =   1
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account No"
            Height          =   195
            Left            =   7920
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   120
            X2              =   10680
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label6 
            Caption         =   "Contact Designation"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Cheque No"
            Height          =   255
            Left            =   5880
            TabIndex        =   14
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lbPolicyNo 
            AutoSize        =   -1  'True
            Caption         =   "Bank No"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lbNames 
            Caption         =   "Contact Names"
            Height          =   255
            Left            =   1800
            TabIndex        =   12
            Top             =   225
            Width           =   1215
         End
      End
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
         Height          =   3855
         Left            =   8040
         TabIndex        =   5
         Top             =   1560
         Width           =   2895
         Begin VB.CommandButton cmdPrintletter 
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
            Left            =   120
            TabIndex        =   33
            Top             =   3240
            Width           =   2655
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
            TabIndex        =   10
            Top             =   2760
            Width           =   2655
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
            TabIndex        =   9
            Top             =   2235
            Width           =   2655
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
            TabIndex        =   8
            Top             =   1710
            Width           =   2655
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
            TabIndex        =   7
            Top             =   1170
            Width           =   2655
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
            TabIndex        =   4
            Top             =   645
            Width           =   2655
         End
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
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   2655
         End
      End
      Begin MSDataGridLib.DataGrid BeneficiaryGrid 
         Height          =   1215
         Left            =   45
         TabIndex        =   28
         Top             =   5520
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   2143
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
         Left            =   120
         TabIndex        =   41
         Top             =   4200
         Width           =   1095
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
         Left            =   120
         TabIndex        =   39
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "SubJect"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Details"
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
         Left            =   120
         TabIndex        =   21
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Date Prepared"
         Height          =   195
         Left            =   4200
         TabIndex        =   18
         Top             =   4590
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prepared By"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   4590
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmALISMBankAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCode As ADODB.Recordset, strcode, iReferenceNo As String, bAddNew As Boolean

Sub ClearControls()
On Error GoTo err
With frmALISMBankAccount
    .txtBankNo.Text = ""
    .txtcontactNames.Text = ""
    .txtDetails.Text = ""
    .txtDatePrepared.Text = ""
    .txtPreparedBy.Text = ""
    .txtSignatoryNames.Text = ""
    .txtSignatorytittle.Text = ""
    .txtContactDesignation.Text = ""
    .cboChequeNo.Text = ""
    .txtAccountNo.Text = ""
    .txtsubject.Text = ""
    .txttopnote.Text = ""
    .txtbottomnote.Text = ""
    
End With
    Exit Sub

err:
ErrorMessage
End Sub

Sub EnableControls()
On Error GoTo err
With frmALISMBankAccount
    .txtBankNo.Locked = False
    .txtcontactNames.Locked = False
    .txtDetails.Locked = False
    .txtDatePrepared.Locked = False
    .txtPreparedBy.Locked = False
    .txtSignatoryNames.Locked = False
    .txtSignatorytittle.Locked = False
    .txtContactDesignation.Locked = False
    .cboChequeNo.Locked = False
    .txtAccountNo.Locked = False
    .txtsubject.Locked = False
    .txttopnote.Locked = False
    .txtbottomnote.Locked = False
    
End With
     Exit Sub

err:
ErrorMessage
End Sub

Sub DisableControls()
On Error GoTo err

   With frmALISMBankAccount
    .txtBankNo.Locked = True
    .txtcontactNames.Locked = True
    .txtDetails.Locked = True
    .txtDatePrepared.Locked = True
    .txtPreparedBy.Locked = True
    .txtSignatoryNames.Locked = True
    .txtSignatorytittle.Locked = True
    .txtContactDesignation.Locked = True
    .cboChequeNo.Locked = True
    .txtAccountNo.Locked = True
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
     txtBankNo.Text = !BankNo & ""
    txtcontactNames = !Contactname & ""
    txtDetails.Text = !Details & ""
    txtDatePrepared = !dateprepared & ""
    txtPreparedBy = !Preparedby & ""
    txtSignatoryNames = !Signatorynames & ""
    txtSignatorytittle = !SignatoryTittle & ""
    txtContactDesignation = !ContactDesignation & ""
    cboChequeNo = !Chequeno & ""
    txtAccountNo = !AccountNo & ""
    txtsubject = !Subject & ""
    txttopnote = !TopNote & ""
    txtbottomnote = !BottomNote & ""
End With
    
    'Call LoadClaimDetails
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

Private Sub cmdAddNew_Click()
On Error GoTo Myerr
        bAddNew = True
        ClearControls
        EnableControls
        DisableCommandButtons
        'LoadClaimDetails
        txtBankNo.FontBold = True
        'txtBankNo.Locked = True
        
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

If txtBankNo.Text = "" Then
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
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
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

                        strQRE = InputBox("Enter Bank No  to search.", "Search Value")
    
                        rsFind.Open "SELECT * FROM ALISPBankAccount WHERE BankNo = '" & strQRE & "';", cnALIS, adOpenKeyset, adLockOptimistic

                        With rsFind
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                Else
                                    txtBankNo.Text = !BankNo
                                    txtcontactNames = !Contactname
                                    txtDetails.Text = !Details
                                    txtDatePrepared = !dateprepared
                                    txtPreparedBy = !Preparedby
                                    txtSignatoryNames = !Signatorynames
                                    txtSignatorytittle = !SignatoryTittle
                                    txtContactDesignation = !ContactDesignation
                                    cboChequeNo = !Chequeno
                                    txtAccountNo = !AccountNo
                                    txtsubject = !Subject
                                    txttopnote = !TopNote
                                    txtbottomnote = !BottomNote
                                    Edit = True
                                End If
                               ' Call LoadClaimDetails
 
                        End With
                        
                        
                        If Edit Then
                                cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ALISPBankAccount WHERE BankNo = '" & txtBankNo.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                                     !BankNo = txtBankNo.Text
                                     !Contactname = txtcontactNames.Text
                                     !Details = txtDetails.Text
                                     !dateprepared = txtDatePrepared.Text
                                     !Preparedby = txtPreparedBy.Text
                                     !Signatorynames = txtSignatoryNames.Text
                                     !SignatoryTittle = txtSignatorytittle.Text
                                     !ContactDesignation = txtContactDesignation.Text
                                     !Chequeno = cboChequeNo.Text
                                     !AccountNo = txtAccountNo.Text
                                      !Subject = txtsubject.Text
                                     !TopNote = txttopnote.Text
                                     !BottomNote = txtbottomnote.Text
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

'On Error GoTo err
        If txtBankNo.Text = "" Then
                MsgBox "BankNo is Required"
                txtBankNo.SetFocus
        
        ElseIf txtcontactNames.Text = "" Then
                MsgBox "The Names is  required"
                txtcontactNames.SetFocus
        
        ElseIf txtDetails.Text = "" Then
                MsgBox "The Details is  required"
                txtDetails.SetFocus
      
        ElseIf txtAccountNo.Text = "" Then
                MsgBox "The AccountNo is  required"
                txtAccountNo.SetFocus
           
           ElseIf txtSignatoryNames.Text = "" Then
                MsgBox "The Signatory Names is  required"
                txtSignatoryNames.SetFocus
            
        ElseIf txtSignatorytittle.Text = "" Then
                MsgBox "The Signatory Tittle is  required"
                txtSignatorytittle.SetFocus
                
                ElseIf cboChequeNo.Text = "" Then
                MsgBox "The Cheque No is  required"
                cboChequeNo.SetFocus
                
                ElseIf txtContactDesignation.Text = "" Then
                MsgBox "The Contact Designation is  required"
                txtAccountNo.SetFocus
                
            
        Else
                GoTo save
        End If
save:
            With RsCode
                    .AddNew
                                     !BankNo = txtBankNo.Text
                                     !Contactname = txtcontactNames.Text
                                     !Details = txtDetails.Text
                                     !Signatorynames = txtSignatoryNames.Text
                                     !SignatoryTittle = txtSignatorytittle.Text
                                     !ContactDesignation = txtContactDesignation.Text
                                     !Chequeno = cboChequeNo.Text
                                     !AccountNo = txtAccountNo.Text
                                     !dateprepared = Date
                                     !Preparedby = CurrentUserName
                                     If txtsubject.Text = "" Then
                                        !Subject = "RE: CHEQUE ISSUED SCHEDULE A/C NO"
                                     Else
                                        !Subject = txtsubject.Text
                                     End If
                                     If txttopnote = "" Then
                                        !TopNote = " Please we have issued the following cheques for payment:"
                                     Else
                                        !TopNote = txttopnote.Text
                                     End If
                                     If txtbottomnote.Text = "" Then
                                         !BottomNote = "Also note that the Cheque payment list can be signed by ONE signatory"
                                     Else
                                        !BottomNote = txtbottomnote.Text
                                     End If
                    .Update
            End With
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
Load frmchequeissueschedule
frmchequeissueschedule.Show 1, Me

Exit Sub
err:
ErrorMessage
End Sub

Private Sub cmdUpdate_Click()


        ValidateData
        Call LoadGrid
        EnableCommandButtons
        DisableControls
        'cmdaddNew.SetFocus


End Sub

Private Sub cmdSearch_Click()

On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
        
        strQRE = InputBox("Enter BankNo to search.", "Search Value")
        rsFind.Open "SELECT * FROM ALISPBankAccount  WHERE BankNo = '" & strQRE & "';", cnALIS, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                            ClearControls
                            txtBankNo.Text = !BankNo
                                    txtcontactNames = !Contactname
                                    txtDetails.Text = !Details
                                    txtDatePrepared = !dateprepared
                                    txtPreparedBy = !Preparedby
                                    txtSignatoryNames = !Signatorynames
                                    txtSignatorytittle = !SignatoryTittle
                                    txtContactDesignation = !ContactDesignation
                                    cboChequeNo = !Chequeno
                                    txtAccountNo = !AccountNo
                                    txtsubject.Text = !Subject & ""
                                    txttopnote.Text = !TopNote & ""
                                    txtbottomnote.Text = !BottomNote & ""
                                    
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
            strcode = "SELECT * from ALISPBankAccount"

   RsCode.Open strcode, cnALIS, adOpenKeyset, adLockOptimistic
    
    'With RsCode
            'If .BOF Or .EOF Then GoTo test:
              '  ShowCode
   ' End With

'test:
    
    DisableControls
    
    Call LoadGrid

    cmdUpdate.Enabled = False

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

   rsGRID.Open "SELECT * FROM ALISPBankAccount", cnALIS, adOpenKeyset, adLockOptimistic
   Set BeneficiaryGrid.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub txtBankNo_LostFocus()
On Error GoTo err
    Dim rscontact As ADODB.Recordset
    Set rscontact = New ADODB.Recordset
    rscontact.Open "Select * from ALISPBank where BankNo= '" & frmALISMBankAccount.txtBankNo & "';", cnALIS, adOpenKeyset, adLockOptimistic
    txtcontactNames.Text = rscontact!Contactname
    txtDatePrepared.Text = Date
    Set rscontact = Nothing
    Exit Sub
err:
    ErrorMessage

End Sub
