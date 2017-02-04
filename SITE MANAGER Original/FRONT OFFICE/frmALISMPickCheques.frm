VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISMPickCheques 
   Caption         =   "Pick Cheques(Recieved)"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   8445
   Icon            =   "frmALISMPickCheques.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   8445
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.Frame Frame3 
         Caption         =   "Cheque Schedule"
         Height          =   2055
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   5535
         Begin MSComctlLib.ListView ListView1 
            Height          =   1695
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   5295
            _ExtentX        =   9340
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
      Begin VB.CommandButton cmdFirstCode 
         Height          =   375
         Index           =   3
         Left            =   4725
         Picture         =   "frmALISMPickCheques.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4320
         Width           =   1815
      End
      Begin VB.CommandButton cmdFirstCode 
         Height          =   375
         Index           =   2
         Left            =   3270
         Picture         =   "frmALISMPickCheques.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdFirstCode 
         Height          =   375
         Index           =   1
         Left            =   1695
         Picture         =   "frmALISMPickCheques.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4320
         Width           =   1575
      End
      Begin VB.CommandButton cmdFirstCode 
         Height          =   375
         Index           =   0
         Left            =   0
         Picture         =   "frmALISMPickCheques.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4320
         Width           =   1695
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
         Left            =   4200
         TabIndex        =   14
         Top             =   3510
         Width           =   1335
      End
      Begin VB.TextBox txtPreparedBy 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1080
         TabIndex        =   12
         Top             =   3510
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   50
         TabIndex        =   9
         Top             =   -120
         Width           =   10815
         Begin VB.ComboBox txtAccountNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   4920
            TabIndex        =   27
            Top             =   960
            Width           =   2655
         End
         Begin VB.ComboBox cbobankNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            TabIndex        =   26
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
            TabIndex        =   16
            Top             =   960
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
            Height          =   375
            Left            =   1800
            TabIndex        =   1
            Top             =   360
            Width           =   5775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account No"
            Height          =   195
            Left            =   3960
            TabIndex        =   28
            Top             =   960
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
            TabIndex        =   17
            Top             =   960
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
            Left            =   1800
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
         Height          =   2895
         Left            =   5760
         TabIndex        =   3
         Top             =   1320
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
            Left            =   720
            TabIndex        =   23
            Top             =   2280
            Width           =   1215
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
            Left            =   1320
            TabIndex        =   8
            Top             =   1560
            Width           =   1215
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
            TabIndex        =   7
            Top             =   1560
            Width           =   1215
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
            TabIndex        =   6
            Top             =   840
            Width           =   1215
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
            Left            =   1320
            TabIndex        =   5
            Top             =   840
            Width           =   1215
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
            Left            =   1320
            TabIndex        =   2
            Top             =   240
            Width           =   1215
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
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid BeneficiaryGrid 
         Height          =   1095
         Left            =   45
         TabIndex        =   18
         Top             =   4800
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1931
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Date Prepared"
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   3510
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prepared By"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   3510
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmALISMPickCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCode As ADODB.Recordset, strcode, iReferenceNo As String, bAddNew As Boolean

Sub ClearControls()
On Error GoTo err
With frmALISMPickCheques
    .txtcontactNames.Text = ""
    .txtDatePrepared.Text = ""
    .txtPreparedBy.Text = ""
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

Sub ShowCode()

On Error GoTo err
    With RsCode
     cboBankNo.Text = !BankNo & ""
    txtDatePrepared = !dateprepared & ""
    txtPreparedBy = !Preparedby & ""
    
   
    End With
    Call cbobankNo_LostFocus
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
    cmdSearch.Enabled = True
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

Private Sub cbobankNo_LostFocus()
On Error GoTo err

Call GetCheques
txtDatePrepared.Text = Date
Call loadCONTACTS
Exit Sub
err:
ErrorMessage
End Sub
Public Sub loadCONTACTS()
On Error GoTo err
    Dim rsCONTACT As ADODB.Recordset
    Set rsCONTACT = New ADODB.Recordset
    rsCONTACT.Open "Select * from ALISPBank where BankNo= '" & frmALISMPickCheques.cboBankNo & "';", cnALIS, adOpenKeyset, adLockOptimistic
    txtcontactNames.Text = rsCONTACT!Contactname
   
    Set rsCONTACT = Nothing
    Exit Sub
err:
    ErrorMessage

End Sub
Public Sub GetCheques()
        On Error GoTo err
    With frmALISMPickCheques
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ChequeNo", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Payee", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "BankNo", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Cheque Date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                 rsLIST.Open "SELECT * from ALISMReceipt where Banked IS NULL and Paymentmethod<>'" & "3" & "';", cnALIS, adOpenKeyset, adLockOptimistic
                 Dim rspersonal As ADODB.Recordset
                 Dim rscorporate As ADODB.Recordset
                 Set rspersonal = New ADODB.Recordset
                 Set rscorporate = New ADODB.Recordset
Dim DF As Integer

                DF = rsLIST.RecordCount
                
                Dim MyList As ListItem
                   
            While Not rsLIST.EOF
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                    If Not IsNull(rsLIST!Payer) Then
                        MyList.SubItems(1) = CStr(rsLIST!Payer)
                    End If
                    If Not IsNull(rsLIST!ReceiptAmount) Then
                        MyList.SubItems(2) = CStr(rsLIST!ReceiptAmount)
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
On Error GoTo Myerr
        bAddNew = True
        ClearControls
        EnableControls
        DisableCommandButtons
        cboBankNo.FontBold = True
Call GetCheques
txtDatePrepared.Text = Date
Call loadCONTACTS
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

If cboBankNo.Text = "" Then
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

Private Sub cmdEdit_Click()

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
                                    cboBankNo.Text = !BankNo
                                    txtcontactNames = !Contactname
                                    txtDatePrepared = !dateprepared
                                    txtPreparedBy = !Preparedby
                                    
                                    
                                    Edit = True
                                End If
    
                       End With
                      
                       If Edit Then
                                cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ALISPBankAccount WHERE BankNo = '" & cboBankNo.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                                     !BankNo = cboBankNo.Text
                                     !Contactname = txtcontactNames.Text
                                     !dateprepared = txtDatePrepared.Text
                                     !Preparedby = txtPreparedBy.Text
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
Dim rsSAVEINQUIRYCFG As ADODB.Recordset
Set rsSAVEINQUIRYCFG = New ADODB.Recordset
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
                GoTo save
        End If
save:
    rsSAVEINQUIRYCFG.Open "Select * from ALISMChequeSchedule", cnALIS, adOpenKeyset, adLockOptimistic
    
                                     
                                     
        Dim i As Integer, j As Integer, k As Variant

        i = Me.ListView1.ListItems.Count: j = 0
        
        For j = 1 To i
            If Me.ListView1.ListItems.Item(j).Checked = True Then
                With rsSAVEINQUIRYCFG
                        .AddNew
                         !ChequeNo = Me.ListView1.ListItems(j).Text
                        
                         !BankNo = cboBankNo.Text
                         !reference = txtReference.Text
                         !dateprepared = Date
                         !Preparedby = txtPreparedBy.Text
                         .Update
                End With
            End If
        Next j
      'Mark the Cheques as Banked in the Receipt File
Dim rsreceipt As ADODB.Recordset
Set rsreceipt = New ADODB.Recordset
Dim l As Integer
rsreceipt.Open "Select * from ALISMReceipt where PaymentMethod<>'" & "3" & "';", cnALIS, adOpenKeyset, adLockOptimistic
 
 For j = 1 To i
    If Me.ListView1.ListItems.Item(j).Checked = True Then
            With rsreceipt
                   For l = 1 To rsreceipt.RecordCount
              If Not IsNull(!ChequeNo) Or Not IsEmpty(!ChequeNo) Then
              
                If Me.ListView1.ListItems(j).Text = !ChequeNo Then
                    !Banked = "Y"
                    .Update
                End If
              End If
                    .MoveNext
                     Next l
           End With
    End If
Next j
cmdAddNew.Enabled = False
        Exit Sub
err:
            If err.Number = 3021 Then Resume Next

            UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVEINQUIRYCFG.CancelUpdate
            rsSAVEINQUIRYCFG.Requery
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
                                    cboBankNo.Text = !BankNo
                                    txtcontactNames = !Contactname
                                    txtDatePrepared = !dateprepared
                                    txtPreparedBy = !Preparedby
                                    
                                    txtAccountNo = !AccountNo
                                    
                            Edit = True
                End If
                            

        End With
        
        Exit Sub

Myerr:
            ErrorMessage

End Sub



Private Sub Form_Load()

On Error GoTo Myerr

    Call OpenConnection
      
            Set RsCode = New Recordset
            strcode = "SELECT * from ALISMChequeSchedule"

   RsCode.Open strcode, cnALIS, adOpenKeyset, adLockOptimistic
    

    frmALISMPickCheques.txtReference = frmALISMBankAccountCfg.txtReferenceNo
    frmALISMPickCheques.txtAccountNo.Text = frmALISMBankAccountCfg.txtAccountNo.Text
    frmALISMPickCheques.cboBankNo.Text = frmALISMBankAccountCfg.cboBankNo.Text
    
    frmALISMPickCheques.txtAccountNo.Locked = True
    frmALISMPickCheques.txtReference.Locked = True
    DisableControls
    
    Call LoadGrid

    cmdUpdate.Enabled = False
On Error Resume Next
Me.Top = 1400
Me.Left = 1800
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

   rsGRID.Open "SELECT ALISMChequeSchedule.ChequeNo,ALISMChequeSchedule.BankNo,ALISMChequeSchedule.ChequeDate,ALISMChequeSchedule.Payee,ALISMChequeSchedule.Amount,ALISMChequeSchedule.Reference FROM ALISMChequeSchedule,ALISMBankAccountCfg where ALISMChequeSchedule.Reference=ALISMBankAccountCfg.Reference and ALISMBankAccountCfg.ReceivedPayed='" & "Received" & "';", cnALIS, adOpenKeyset, adLockOptimistic
   Set BeneficiaryGrid.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub txtBankNo_LostFocus()
On Error GoTo err
    Dim rsCONTACT As ADODB.Recordset
    Set rsCONTACT = New ADODB.Recordset
    rsCONTACT.Open "Select * from ALISPBank where BankNo= '" & frmALISMPickCheques.cboBankNo & "';", cnALIS, adOpenKeyset, adLockOptimistic
    txtcontactNames.Text = rsCONTACT!Contactname
    txtDatePrepared.Text = Date
    Set rsCONTACT = Nothing
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
rsloadbankacc.Open "Select * from ALISPBankAccount where BankNo= ' " & frmALISMPickCheques.cboBankNo & "';", cnALIS, adOpenKeyset, adLockOptimistic
For i = 1 To rsloadbankacc.RecordCount
    txtAccountNo.AddItem rsloadbankacc!AccountNo
    rsloadbankacc.MoveNext
Next i
Set rsloadbankacc = Nothing

Exit Sub
err:
ErrorMessage
End Sub
