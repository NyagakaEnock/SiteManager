VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISMPickChequesPay 
   Caption         =   "Pick Cheques (Pay)"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10035
   Icon            =   "frmALISMPickChequesPay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10035
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.TextBox txtChequeNo 
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
         Left            =   7440
         TabIndex        =   22
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cheques Scheduled"
         Height          =   2535
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   7095
         Begin MSDataGridLib.DataGrid BeneficiaryGrid 
            Height          =   2175
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   3836
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
      Begin VB.Frame Frame3 
         Caption         =   "Cheque Not Scheduled"
         Height          =   2055
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   7095
         Begin MSComctlLib.ListView ListView1 
            Height          =   1695
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   2990
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
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   50
         TabIndex        =   9
         Top             =   -120
         Width           =   10815
         Begin VB.ComboBox txtAccountNo 
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
            Left            =   6120
            TabIndex        =   18
            Top             =   960
            Width           =   3615
         End
         Begin VB.ComboBox cbobankNo 
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
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtReference 
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
            TabIndex        =   12
            Top             =   930
            Width           =   2055
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
            Height          =   360
            Left            =   1800
            TabIndex        =   1
            Top             =   360
            Width           =   7935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account No"
            Height          =   195
            Left            =   5160
            TabIndex        =   19
            Top             =   1020
            Width           =   855
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   120
            X2              =   10680
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label6 
            Caption         =   "Reference No"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   990
            Width           =   1455
         End
         Begin VB.Label lbPolicyNo 
            AutoSize        =   -1  'True
            Caption         =   "Bank No"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lbNames 
            Caption         =   "Contact Names"
            Height          =   255
            Left            =   4800
            TabIndex        =   10
            Top             =   105
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
         Height          =   3735
         Left            =   7320
         TabIndex        =   3
         Top             =   2280
         Width           =   2655
         Begin VB.CommandButton cmdPrintletter 
            Caption         =   "&Print "
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
            TabIndex        =   14
            Top             =   3120
            Width           =   2295
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
            TabIndex        =   8
            Top             =   2640
            Width           =   2295
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
            TabIndex        =   7
            Top             =   2160
            Width           =   2295
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
            TabIndex        =   6
            Top             =   1200
            Width           =   2295
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
            TabIndex        =   5
            Top             =   1680
            Width           =   2295
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
            TabIndex        =   2
            Top             =   720
            Width           =   2295
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
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   2295
         End
      End
   End
End
Attribute VB_Name = "frmALISMPickChequesPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCode As adodb.Recordset, strcode, iReferenceNo As String, bAddNew As Boolean

Sub ClearControls()
On Error GoTo err
        With frmALISMPickCheques
            .txtcontactNames.Text = ""
            .ListView1.ListItems.Clear
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Sub EnableControls()
On Error GoTo err
    With frmALISMPickCheques
        .cboBankNo.Locked = False
        .txtDatePrepared.Locked = False
        .txtPreparedBy.Locked = False
        .txtAccountNo.Locked = False
        
    End With
Exit Sub

err:
ErrorMessage
End Sub

Sub DisableControls()
On Error GoTo err
   With frmALISMPickCheques
        .cboBankNo.Locked = True
        .txtcontactNames.Locked = True
        .txtDatePrepared.Locked = True
        .txtAccountNo.Locked = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub ShowCode()
On Error GoTo err
    With RsCode
        frmALISMPickChequesPay.cboBankNo.Text = !BankNo & ""
    End With
    loadDETAILS

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cbobankNo_GotFocus()
On Error GoTo err
    
    Screen.ActiveForm.cboBankNo.Clear
    
    Dim i As Integer
    Dim rsloadbank As adodb.Recordset
    Set rsloadbank = New adodb.Recordset
    
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
Private Sub loadDETAILS()
On Error GoTo err
cboBankNo.Clear
    Dim i As Integer
    Dim rsloadbank As adodb.Recordset
    Set rsloadbank = New adodb.Recordset
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

Private Sub cbobankNo_LostFocus()
    GetCheques
    Screen.ActiveForm.txtDatePrepared.Text = Date
    loadCONTACTS
End Sub

Public Sub loadCONTACTS()
On Error GoTo err

    Dim rsCONTACT As adodb.Recordset
    Set rsCONTACT = New adodb.Recordset
    
    rsCONTACT.Open "Select * from ALISPBank where BankNo= '" & frmALISMPickChequesPay.cboBankNo & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsCONTACT
        If .EOF Or .BOF Then Set rsCONTACT = Nothing: Exit Sub
            Screen.ActiveForm.txtcontactNames.Text = rsCONTACT!Contactname
    End With
         
    
rsCONTACT.Close
    
Exit Sub
err:
    ErrorMessage

End Sub
Public Sub GetCheques()
On Error GoTo err
        
        With frmALISMPickChequesPay
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ChequeNo", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Payee", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "BankNo", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Cheque Date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As adodb.Recordset
                Set rsLIST = New adodb.Recordset
                'rsLIST.Open "SELECT ALISMCheque.ChequeNo,ALISMCheque.Amount,ALISMCheque.Chequedate,ALISMCheque.BankNo,ALISMCheque.PayeeDetails from ALISMCheque where Status = 'CHECK ISSUANCE' and ALISMCheque.BankNo= '" & frmALISMPickChequesPay.cbobankNo & "';", cnALIS, adOpenKeyset, adLockOptimistic
                strSQL = ""
                strSQL = "SELECT ALISMCheque.ChequeNo,ALISMCheque.Amount,ALISMCheque.Chequedate,ALISMCheque.BankNo,ALISMCheque.PayeeDetails from ALISMCheque where ALISMCheque.BankNo= '" & frmALISMPickChequesPay.cboBankNo & "'and scheduled is null and ALISMCheque.Authorized = 'Y' ;"
                rsLIST.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
                
                Dim rspersonal As adodb.Recordset
                Dim rscorporate As adodb.Recordset
                Set rspersonal = New adodb.Recordset
                Set rscorporate = New adodb.Recordset
                Dim DF As Integer

                DF = rsLIST.RecordCount
                
                Dim MyList As ListItem
                   
            While Not rsLIST.EOF
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                    If Not IsNull(rsLIST!Payeedetails) Then
                        MyList.SubItems(1) = CStr(rsLIST!Payeedetails)
                    End If
                    If Not IsNull(rsLIST!Amount) Then
                        MyList.SubItems(2) = CStr(rsLIST!Amount)
                    End If
                     If Not IsNull(rsLIST!BankNo) Then
                        MyList.SubItems(3) = CStr(rsLIST!BankNo)
                    End If
                    If Not IsNull(rsLIST!ChequeDate) Then
                        MyList.SubItems(4) = CStr(rsLIST!ChequeDate)
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
Private Sub cmdAddNew_Click()
        bAddNew = True
        ClearControls
        EnableControls
        disableButtons
        GetCheques
        loadCONTACTS
End Sub


Private Sub cmdCancel_Click()
        enableButtons
        ClearControls
        DisableControls
End Sub


Sub loadDEFAULTS()

On Error GoTo err

    Dim rsDEFA As adodb.Recordset, strDEFA, LSetupCode As String
    Set rsDEFA = New Recordset
    
    LSetupCode = 0
    
    strDEFA = "Select * from ALISMReference where referenceNo LIKE  '" & iReferenceNo & "';"
    rsDEFA.Open strDEFA, cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsDEFA
        If .EOF And .BOF Then Exit Sub
            Screen.ActiveForm.txtcontactNames = Trim(!othernames) + " " + Trim(!Surname)
            Screen.ActiveForm.txtDatePrepared = Date
   End With

rsDEFA.Close
strDEFA = ""

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub ValidateData()
On Error GoTo err
        
        If cboBankNo.Text = "" Then
                MsgBox "BankNo is Required"
                cboBankNo.SetFocus
        ElseIf txtReference.Text = "" Then
                MsgBox "The Names is  required"
                txtcontactNames.SetFocus
        ElseIf txtAccountNo.Text = "" Then
                MsgBox "The AccountNo is  required"
                txtAccountNo.SetFocus
        Else
                bsaveRECORD = True
        End If
        
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub saveRECORD()
On Error GoTo err
        
        Dim rsSAVE As adodb.Recordset
        Set rsSAVE = New adodb.Recordset
        
        strSQL = ""
        strSQL = "Select * from ALISMCheque where Scheduled is null and bankNo = '" & frmALISMPickChequesPay.cboBankNo.Text & "' and chequeNo = '" & frmALISMPickChequesPay.txtChequeNo.Text & "'  ;"
        rsSAVE.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
                        
        With rsSAVE
                                
                  If .BOF Or .EOF = True Then Exit Sub
                              
                        !reference = frmALISMPickChequesPay.txtReference.Text
                        !Status = "CHK-SCHEDULED"
                        !sCHEDULED = "Y"
                        !DateScheduled = Date
                        !ScheduledBy = CurrentUserName
                        .Update
        End With
 
Exit Sub
err:
        If err.Number = 3021 Then Resume Next

        UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
        End If
  
End Sub


Private Sub cmdPrintletter_Click()
    Load frmchequeissueschedulePay
    frmchequeissueschedulePay.Show 1, Me
End Sub

Private Sub cmdUpdate_Click()
        bsaveRECORD = False
        ValidateData
        If bsaveRECORD = True Then
            GetCheckedBoxes
            LoadGrid
            bsaveRECORD = False
        End If
        enableButtons
        DisableControls
        GetCheques
        
End Sub

Private Sub cmdSearch_Click()

On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As adodb.Recordset, Edit As Boolean

        Set rsFind = New adodb.Recordset
        
        strQRE = InputBox("Enter BankNo to search.", "Search Value")
        rsFind.Open "SELECT * FROM ALISPBankAccount  WHERE BankNo = '" & strQRE & "';", cnALIS, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                                    ClearControls
                                    cboBankNo.Text = !BankNo
                                    txtcontactNames = !Contactname
                                    txtAccountNo = !AccountNo
                                    
                            Edit = True
                End If
        End With
        
    Exit Sub

Myerr:
      ErrorMessage

End Sub

Private Sub Form_Load()
    OpenConnection
    DisableControls
    LoadGrid
    enableButtons
    LoadRECORD
End Sub
Private Sub LoadRECORD()
    frmALISMPickChequesPay.txtReference = frmALISMBankAccountCfgPay.txtReferenceNo.Text
    frmALISMPickChequesPay.txtAccountNo.Text = frmALISMBankAccountCfgPay.txtAccountNo.Text
    frmALISMPickChequesPay.cboBankNo.Text = frmALISMBankAccountCfgPay.cboBankNo.Text
End Sub
Private Sub GetCheckedBoxes()
   
            Dim j, i As Integer, strINSTALLMENT As String
       
            j = ListView1.ListItems.Count
            
            For i = 1 To j
                    If ListView1.ListItems(i).Checked = True Then
                        strINSTALLMENT = ListView1.ListItems(i).Text
                        frmALISMPickChequesPay.txtChequeNo.Text = strINSTALLMENT
                        
                        saveRECORD
                    End If
                    strINSTALLMENT = ""
            Next i
            
        

End Sub

Private Sub updateALLRECORD()
    Dim rsSAVE As adodb.Recordset
    strSQL = ""
    Set rsSAVE = New adodb.Recordset
    strSQL = "SELECT * FROM ALISMInstallment WHERE PolicyNo = '" & Screen.ActiveForm.cboPolicyNo.Text & "' and Installment = '" & Screen.ActiveForm.txtInstallment.Text & "' ;  "
    rsSAVE.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
  


End Sub

Private Sub LoadGrid()
On Error GoTo err

    Dim rsGRID As adodb.Recordset
    Set rsGRID = New Recordset
    
    strSQL = ""
    strSQL = "SELECT * FROM ALISMCheque where ALISMCheque.Reference= '" & frmALISMPickChequesPay.txtReference.Text & "';"
    
    rsGRID.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
    Set BeneficiaryGrid.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Text1_Change()

End Sub
