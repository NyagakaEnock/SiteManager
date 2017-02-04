VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISMBankAccountCfgPay 
   Caption         =   "Cheque Issue Schedule Letter (Pay)"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   11880
   Icon            =   "frmALISMBankAccountCfgPay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11775
      Begin TabDlg.SSTab SSTabBankCheques 
         Height          =   7335
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   12938
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Bank Cheques"
         TabPicture(0)   =   "frmALISMBankAccountCfgPay.frx":0442
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame3"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "List Letters "
         TabPicture(1)   =   "frmALISMBankAccountCfgPay.frx":045E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "BeneficiaryGrid"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "frmALISMBankAccountCfgPay.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.Frame Frame3 
            Height          =   2775
            Left            =   120
            TabIndex        =   31
            Top             =   3960
            Width           =   8535
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   2415
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   8295
               _ExtentX        =   14631
               _ExtentY        =   4260
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
         Begin VB.Frame Frame4 
            Height          =   1215
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   11295
            Begin VB.TextBox txtBankName 
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
               Left            =   3120
               TabIndex        =   30
               Top             =   240
               Width           =   4455
            End
            Begin VB.TextBox txtAccountNo 
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
               Left            =   8760
               TabIndex        =   29
               Top             =   240
               Width           =   2295
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
               TabIndex        =   22
               Top             =   1920
               Width           =   2655
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
               TabIndex        =   1
               Top             =   240
               Width           =   1935
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
               Height          =   375
               Left            =   8760
               TabIndex        =   21
               Top             =   720
               Width           =   2295
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
               Left            =   4440
               TabIndex        =   20
               Top             =   720
               Width           =   3135
            End
            Begin VB.TextBox txtReferenceNo 
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
               Left            =   1200
               TabIndex        =   2
               Top             =   720
               Width           =   1935
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Account No"
               Height          =   195
               Left            =   7680
               TabIndex        =   27
               Top             =   330
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Telephone No"
               Height          =   255
               Left            =   7680
               TabIndex        =   26
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Bank No"
               Height          =   195
               Left            =   120
               TabIndex        =   25
               Top             =   330
               Width           =   630
            End
            Begin VB.Label lbPolicyNo 
               AutoSize        =   -1  'True
               Caption         =   "Reference "
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   810
               Width           =   795
            End
            Begin VB.Label lbNames 
               Caption         =   "Contact Names"
               Height          =   255
               Left            =   3240
               TabIndex        =   23
               Top             =   780
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
            Height          =   5175
            Left            =   8715
            TabIndex        =   12
            Top             =   1560
            Width           =   2775
            Begin VB.CommandButton cmdPrint 
               Caption         =   "Che&que Schedule"
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
               Top             =   3960
               Width           =   2415
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
               TabIndex        =   0
               Top             =   360
               Width           =   2415
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
               TabIndex        =   18
               Top             =   885
               Width           =   2415
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
               TabIndex        =   17
               Top             =   1395
               Width           =   2415
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
               TabIndex        =   16
               Top             =   1920
               Width           =   2415
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
               TabIndex        =   15
               Top             =   2445
               Width           =   2415
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
               TabIndex        =   14
               Top             =   2955
               Width           =   2415
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
               TabIndex        =   13
               Top             =   3480
               Width           =   2415
            End
         End
         Begin VB.Frame Frame6 
            Height          =   2415
            Left            =   120
            TabIndex        =   8
            Top             =   1560
            Width           =   8535
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
               Left            =   1200
               TabIndex        =   3
               Top             =   240
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
               Height          =   735
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Top             =   720
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
               Height          =   735
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Top             =   1560
               Width           =   7215
            End
            Begin VB.Label Label9 
               Caption         =   "Subject"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   300
               Width           =   975
            End
            Begin VB.Label Label10 
               Caption         =   "Top Note"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label11 
               Caption         =   "Bottom Note"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   1800
               Width           =   1095
            End
         End
         Begin MSDataGridLib.DataGrid BeneficiaryGrid 
            Height          =   6255
            Left            =   -74880
            TabIndex        =   28
            Top             =   360
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   11033
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
Attribute VB_Name = "frmALISMBankAccountCfgPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsLOADGRID As clsALISGRID
Dim RsCode As ADODB.Recordset, strcode, iReferenceNo As String, bAddNew As Boolean

Private Sub ClearControls()
On Error GoTo err
        With frmALISMBankAccountCfgPay
            .cboBankNo.Text = ""
            .txtcontactNames.Text = ""
            .txtsubject.Text = ""
            .txttopnote.Text = ""
            .txtbottomnote.Text = ""
            .txtReferenceNo.Text = ""
            .txttelephoneNo.Text = ""
            .txtBankName.Text = ""
            .txtAccountNo.Text = ""
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub EnableControls()
On Error GoTo err
    With frmALISMBankAccountCfgPay
        .cboBankNo.Locked = False
        .txtcontactNames.Locked = False
        .txtsubject.Locked = False
        .txttopnote.Locked = False
        .txtbottomnote.Locked = False
        .txtAccountNo.Locked = False
        .txtBankName.Locked = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub DisableControls()
On Error GoTo err
   With frmALISMBankAccountCfgPay
        .cboBankNo.Locked = True
        .txtcontactNames.Locked = True
        .txtAccountNo.Locked = True
        .txtsubject.Locked = True
        .txttopnote.Locked = True
        .txtbottomnote.Locked = True
        .txtBankName.Locked = True
    End With

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub loadBankName()
On Error GoTo err
            
          Dim rsbank As ADODB.Recordset, strBANK As String
          Set rsbank = New ADODB.Recordset
          
          strBANK = "SELECT * FROM ALISPbankAccount where bankno = '" & frmALISMBankAccountCfgPay.cboBankNo.Text & "' ;"
          rsbank.Open strBANK, cnALIS, adOpenKeyset, adLockOptimistic

            With rsbank
                    If .EOF Or .BOF Then
                        MsgBox "The Bank Selected is invalid", vbOKOnly
                        Exit Sub
                    End If
                
                    frmALISMBankAccountCfgPay.txtBankName = !details
                    frmALISMBankAccountCfgPay.txtAccountNo = !AccountNo
                    frmALISMBankAccountCfgPay.txtcontactNames.Text = !Contactname
                    frmALISMBankAccountCfgPay.txttelephoneNo.Text = !contactTelephoneNo

            End With
            
        rsbank.Close
        strBANK = ""
Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub ShowCode()

On Error GoTo err
    With RsCode
            frmALISMBankAccountCfgPay.txtReferenceNo.Text = !ReferenceNo
            frmALISMBankAccountCfgPay.cboBankNo.Text = !BankNo & ""
            frmALISMBankAccountCfgPay.txtAccountNo = !AccountNo & ""
            frmALISMBankAccountCfgPay.txtsubject = !Subject & ""
            frmALISMBankAccountCfgPay.txttopnote = !TopNote & ""
            frmALISMBankAccountCfgPay.txtbottomnote = !BottomNote & ""
      End With

Exit Sub
err:
ErrorMessage

End Sub

Private Sub cbobankNo_GotFocus()
On Error GoTo err
          Dim rsBANKGF As ADODB.Recordset, strBANKGF As String
          Set rsBANKGF = New ADODB.Recordset
          
          strBANKGF = "SELECT * FROM ALISPbankAccount;"
          rsBANKGF.Open strBANKGF, cnALIS, adOpenKeyset, adLockOptimistic

          frmALISMBankAccountCfgPay.cboBankNo.Clear

          With rsBANKGF
                  If .EOF Or .BOF Then Exit Sub
                  Do Until .EOF
                      frmALISMBankAccountCfgPay.cboBankNo.AddItem !details & ""
                      .MoveNext
                  Loop
          End With
        
          rsBANKGF.Close
          strBANKGF = ""
 
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdAddNew_Click()
        bAddNew = True
        ClearControls
        EnableControls
        disableButtons
        LoadDefaults
End Sub


Private Sub cmdCancel_Click()
        enableButtons
        ClearControls
        DisableControls
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Myerr

If frmALISMBankAccountCfgPay.txtReferenceNo.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
ElseIf frmALISMBankAccountCfgPay.txtcontactNames.Text = "" Then
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

                        strQRE = InputBox("Enter Reference No  to search.", "Search Value")
    
                        rsFind.Open "SELECT * FROM ALISMBankAccountCfg WHERE Reference = '" & strQRE & "' and ReceivedPayed='" & "Payed" & "';", cnALIS, adOpenKeyset, adLockOptimistic

                        With rsFind
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                Else
                                    frmALISMBankAccountCfgPay.cboBankNo.Text = !BankNo
                                    frmALISMBankAccountCfgPay.txtReferenceNo.Text = !reference
                                    frmALISMBankAccountCfgPay.txtsubject = !Subject
                                    frmALISMBankAccountCfgPay.txttopnote = !TopNote
                                    frmALISMBankAccountCfgPay.txtbottomnote = !BottomNote
                                    frmALISMBankAccountCfgPay.txtAccountNo.Text = !AccountNo & ""
                                    Edit = True
                                End If
 
                        End With
                                                
                        If Edit Then
                                cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ALISMBankAccountCfg WHERE Reference = '" & frmALISMBankAccountCfgPay.txtReferenceNo.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                                     !BankNo = frmALISMBankAccountCfgPay.cboBankNo.Text
                                     !Subject = frmALISMBankAccountCfgPay.txtsubject.Text
                                     !TopNote = frmALISMBankAccountCfgPay.txttopnote.Text
                                     !BottomNote = frmALISMBankAccountCfgPay.txtbottomnote.Text
                                     !AccountNo = frmALISMBankAccountCfgPay.txtAccountNo.Text
                            .Update
                            .Requery
                            Edit = False
                    End With
                
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


Private Sub ValidateData()
On Error GoTo err
    
    With frmALISMBankAccountCfgPay
            If .cboBankNo.Text = "" Then
                   MsgBox "BankNo is Required"
                   .cboBankNo.SetFocus
        
            ElseIf .txtcontactNames.Text = "" Then
                   MsgBox "The Names is  required"
                   .txtcontactNames.SetFocus
                      
            ElseIf .txtAccountNo.Text = "" Then
                   MsgBox "The Telephone No is  required"
                   .txtAccountNo.SetFocus
            
            ElseIf .txtsubject.Text = "" Then
                    MsgBox "The Subject cannot be left Blank"
                    .txtsubject.SetFocus
            
            ElseIf .txttopnote.Text <= "" Then
                    MsgBox "The Top Note cannot be Left Bank"
                    .txttopnote.SetFocus
            
            ElseIf .txtbottomnote.Text <= "" Then
                    MsgBox "The Bottom Note cannot be Left Blank"
                    .txtbottomnote.SetFocus
            Else
                    bsaveRECORD = True
            End If
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub LoadDefaults()
On Error GoTo err
        
        With frmALISMBankAccountCfgPay
            .txtsubject.Text = "RE: CHEQUE ISSUED SCHEDULE A/C NO"
            .txttopnote = " Please we have issued the following cheques for payment:"
            .txtbottomnote.Text = "Also note that the Cheque payment list can be signed by ONE signatory"
        End With
Exit Sub

err:
    ErrorMessage
End Sub
     
Private Sub saveRECORD()
    Dim rsSAVE As ADODB.Recordset
    Set rsSAVE = New ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT * FROM ALISMBankAccountCfg WHERE Reference =  '" & frmALISMBankAccountCfgPay.txtReferenceNo & " ' ;"
    rsSAVE.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
                      
    With rsSAVE
            .AddNew
             !reference = frmALISMBankAccountCfgPay.txtReferenceNo.Text
             !BankNo = frmALISMBankAccountCfgPay.cboBankNo.Text
             !dateprepared = Date
             !Preparedby = CurrentUserName
             !AccountNo = frmALISMBankAccountCfgPay.txtAccountNo.Text
             !ReceivedPayed = "Payed"
             !Subject = frmALISMBankAccountCfgPay.txtsubject.Text
             !TopNote = frmALISMBankAccountCfgPay.txttopnote.Text
             !BottomNote = frmALISMBankAccountCfgPay.txtbottomnote.Text
              bsaveRECORD = False
            .Update
    End With

rsSAVE.Close
strSQL = ""
            
Exit Sub

err:
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                    rsSAVE.CancelUpdate
                    rsSAVE.Requery
        Else
                    UpdateErrorMessage
        End If
  
End Sub


Private Sub cmdPrintletter_Click()
On Error GoTo err
        If frmALISMBankAccountCfgPay.txtReferenceNo.Text = "" Then
                MsgBox " The reference is required before proceeding", vbCritical
                Exit Sub
        Else
                'Load frmALISMPickChequesPay
                'frmALISMPickChequesPay.Show 1, Me
        End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdUpdate_Click()
        checkChequeNotScheduled
        If bCONTINUE = False Then Exit Sub
        bsaveRECORD = False
        ValidateData
        If bsaveRECORD = True Then
                generateReference
                saveRECORD
                LoadGrid
                enableButtons
                DisableControls
                'Load frmALISMPickChequesPay
                'frmALISMPickChequesPay.Show 1, Me

        End If
End Sub
Public Sub checkChequeNotScheduled()
On Error GoTo err

    bCONTINUE = True
    
    Dim rsUNSCHEDULED As ADODB.Recordset
    Set rsUNSCHEDULED = New ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT * FROM ALISMCheque WHERE BankNo =  '" & frmALISMBankAccountCfgPay.cboBankNo.Text & " ' and (scheduled = '' or scheduled is null );"
    rsUNSCHEDULED.Open strSQL, cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsUNSCHEDULED
            If .BOF Or .EOF Then
                MsgBox "There are No Cheque that have not been Scheduled", vbOKOnly
                bCONTINUE = False
                Exit Sub
            End If
            
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub generateReference()
On Error GoTo err
        Dim rsLAST As ADODB.Recordset, strLAST As String
        
        Set rsLAST = New Recordset
      
        strLAST = "SELECT * FROM ALISPBankAccount Where BankNo = '" & frmALISMBankAccountCfgPay.cboBankNo.Text & "' ;"
        rsLAST.Open strLAST, cnALIS, adOpenKeyset, adLockOptimistic

        With rsLAST
                If .BOF Or .EOF Then Exit Sub
                
                frmALISMBankAccountCfgPay.txtReferenceNo = !reference
                
                Select Case Len(frmALISMBankAccountCfgPay.txtReferenceNo)
                    Case 1: frmALISMBankAccountCfgPay.txtReferenceNo = Trim(!ReferencePrefix) + "/00000" + Trim(frmALISMBankAccountCfgPay.txtReferenceNo)
                    Case 2: frmALISMBankAccountCfgPay.txtReferenceNo = Trim(!ReferencePrefix) + "/0000" + Trim(frmALISMBankAccountCfgPay.txtReferenceNo)
                    Case 3: frmALISMBankAccountCfgPay.txtReferenceNo = Trim(!ReferencePrefix) + "/000" + Trim(frmALISMBankAccountCfgPay.txtReferenceNo)
                    Case 4: frmALISMBankAccountCfgPay.txtReferenceNo = Trim(!ReferencePrefix) + "/00" + Trim(frmALISMBankAccountCfgPay.txtReferenceNo)
                    Case 5: frmALISMBankAccountCfgPay.txtReferenceNo = Trim(!ReferencePrefix) + "/0" + Trim(frmALISMBankAccountCfgPay.txtReferenceNo)
                    Case 6: frmALISMBankAccountCfgPay.txtReferenceNo = Trim(!ReferencePrefix) + "/" + Trim(frmALISMBankAccountCfgPay.txtReferenceNo)
                End Select
                
                !reference = !reference + 1
                .Update
        End With

Exit Sub

err:
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                    rsLAST.CancelUpdate
                    rsLAST.Requery
        Else
                    UpdateErrorMessage
        End If

End Sub
Private Sub cmdSearch_Click()
        ClearControls
        loadRECORD
        
        If bCONTINUE = True Then
                loadDETAILS
                loadBANK
                Set rsLOADGRID = New clsALISGRID
                rsLOADGRID.LoadChequeScheduled
                Set rsLOADGRID = Nothing
        End If
End Sub

Private Sub loadRECORD()
On Error GoTo Myerr
        
        bCONTINUE = False
        
        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean
        
        Set rsFind = New ADODB.Recordset
      
        strQRE = InputBox("Enter ReferenceNo to search.", "Search Value")
        rsFind.Open "SELECT * FROM ALISMBankAccountCfg  WHERE Reference = '" & strQRE & "' and ReceivedPayed='" & "Payed" & "';", cnALIS, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                            frmALISMBankAccountCfgPay.cboBankNo.Text = !BankNo
                            frmALISMBankAccountCfgPay.txtReferenceNo.Text = !reference
                            frmALISMBankAccountCfgPay.txtsubject.Text = !Subject & ""
                            frmALISMBankAccountCfgPay.txttopnote.Text = !TopNote & ""
                            frmALISMBankAccountCfgPay.txtbottomnote.Text = !BottomNote & ""
                            frmALISMBankAccountCfgPay.txtAccountNo.Text = !AccountNo & ""
                            Me.cmdPrintletter.Enabled = True
                            'Call frmALISMBankAccountCfgPay.cbobankNo_LostFocus - replaced by Abwao
                            Edit = True
                End If
                bCONTINUE = True
        End With
        
Exit Sub

Myerr:
            ErrorMessage

End Sub



Private Sub cmdPrint_Click()
If frmALISMBankAccountCfgPay.txtReferenceNo.Text > "" Then
    Load frmchequeissueschedulePay
    frmchequeissueschedulePay.Show 1, Me
End If
End Sub

Private Sub Form_Activate()
    DisableControls
    enableButtons
    loadCHEQUE

End Sub

Private Sub Form_Load()

    OpenConnection
      
    Set RsCode = New Recordset
    strcode = "SELECT * from ALISMBankAccountCfg"

    RsCode.Open strcode, cnALIS, adOpenKeyset, adLockOptimistic
End Sub


Private Sub LoadGrid()
    
On Error GoTo err
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New Recordset

   rsGRID.Open "SELECT * FROM ALISMBankAccountCfg where ReceivedPayed='" & "Payed" & "';", cnALIS, adOpenKeyset, adLockOptimistic
   Set BeneficiaryGrid.DataSource = rsGRID

Exit Sub
err:
    ErrorMessage
End Sub


Private Sub loadDETAILS()
On Error GoTo err
        
        Dim rsBANKLF As ADODB.Recordset
        
        Set rsBANKLF = New ADODB.Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBankAccount WHERE Details = '" & frmALISMBankAccountCfgPay.cboBankNo.Text & "' ", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                frmALISMBankAccountCfgPay.cboBankNo.Text = !BankNo
                frmALISMBankAccountCfgPay.txtBankName.Text = !details & ""
                frmALISMBankAccountCfgPay.txtAccountNo = !AccountNo & ""
                frmALISMBankAccountCfgPay.txtcontactNames.Text = !Contactname & ""
                frmALISMBankAccountCfgPay.txttelephoneNo.Text = !TelephoneNo & ""

        End With
        
        
rsBANKLF.Close

Exit Sub

err:
        ErrorMessage
End Sub

Private Sub cbobankNo_LostFocus()
    If frmALISMBankAccountCfgPay.cboBankNo.Text > "" Then
            loadBANKDETAILS
            loadBANK
            Set rsLOADGRID = New clsALISGRID
            rsLOADGRID.LoadChequeNotScheduled
            Set rsLOADGRID = Nothing
            
            Call LoadGrid

            
    End If
End Sub

Private Sub loadCHEQUE()
'On Error GoTo err

    If frmALISMBankAccountCfgPay.cboBankNo.Text > "" Then
            loadBANKDETAILS
            loadBANK
            Set rsLOADGRID = New clsALISGRID
            rsLOADGRID.LoadChequeNotScheduled
            Set rsLOADGRID = Nothing
            
            Call LoadGrid
    End If
Exit Sub

err:
    ErrorMessage
End Sub

            

Private Sub loadBANK()
On Error GoTo err
        
        Dim rsBANKLF As ADODB.Recordset
        
        Set rsBANKLF = New ADODB.Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBankAccount WHERE BankNo = '" & frmALISMBankAccountCfgPay.cboBankNo.Text & "' ", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                frmALISMBankAccountCfgPay.cboBankNo.Text = !BankNo
                frmALISMBankAccountCfgPay.txtBankName.Text = !details & ""
                frmALISMBankAccountCfgPay.txtAccountNo = !AccountNo & ""
                frmALISMBankAccountCfgPay.txtcontactNames.Text = !Contactname & ""
                frmALISMBankAccountCfgPay.txttelephoneNo.Text = !TelephoneNo & ""
        End With
        
        
rsBANKLF.Close

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub loadBANKDETAILS()
On Error GoTo err
        
        Dim rsBANKLF As ADODB.Recordset
        
        Set rsBANKLF = New ADODB.Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBankAccount WHERE details = '" & frmALISMBankAccountCfgPay.cboBankNo.Text & "' ", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                frmALISMBankAccountCfgPay.cboBankNo.Text = !BankNo
        End With
        
        
rsBANKLF.Close

Exit Sub

err:
        ErrorMessage

End Sub


