VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPPaymentMethod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Methods"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   HelpContextID   =   6
   Icon            =   "frmODASPaymentMethod.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9450
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Frame Frame3 
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
         Left            =   7920
         TabIndex        =   12
         Top             =   840
         Width           =   1215
         Begin VB.CommandButton cmdPrint 
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
            Picture         =   "frmODASPaymentMethod.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3240
            Width           =   975
         End
         Begin VB.CommandButton cmdaddNew 
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
            Picture         =   "frmODASPaymentMethod.frx":040C
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdate 
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
            Picture         =   "frmODASPaymentMethod.frx":050E
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdedit 
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
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1725
            Width           =   975
         End
         Begin VB.CommandButton cmdsearch 
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
            Picture         =   "frmODASPaymentMethod.frx":0610
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2220
            Width           =   975
         End
         Begin VB.CommandButton cmddelete 
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
            Picture         =   "frmODASPaymentMethod.frx":0712
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2715
            Width           =   975
         End
         Begin VB.CommandButton cmdCancel 
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
            Picture         =   "frmODASPaymentMethod.frx":0814
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1230
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   7695
         Begin VB.OptionButton optRequireAccountDetails 
            Caption         =   "Require Details?"
            Height          =   255
            Left            =   5640
            TabIndex        =   11
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optJournal 
            Caption         =   "Journal?"
            Height          =   255
            Left            =   2880
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optCash 
            Caption         =   "Cash?"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3135
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   7695
         Begin MSComctlLib.ListView ListView1 
            Height          =   2775
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   7455
            _ExtentX        =   13150
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
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9015
         Begin VB.TextBox txtPaymentMethod 
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
            Left            =   1440
            TabIndex        =   3
            Top             =   240
            Width           =   1320
         End
         Begin VB.TextBox txtPaymentMethodDescription 
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
            Left            =   4080
            TabIndex        =   2
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Payment Method"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   323
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Left            =   2880
            TabIndex        =   4
            Top             =   323
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frmODASPPaymentMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rep, bVAL As Boolean
Dim strPAYMENT As String
Dim rsPAYMENT As ADODB.Recordset

Private Sub cmdAddNew_Click()
    rep = True
    enableALLRECORD
    clearALLRECORD
    disableButtons
End Sub

Private Sub loadRECORD()
On Error GoTo err
    
    strSQL = "select * from ODASPPaymentMethod where PaymentMethod = '" & frmODASPPaymentMethod.txtPaymentMethod.Text & "';"
    Set rsCONTROL = New ADODB.Recordset
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsCONTROL
        If .EOF And .BOF Then Exit Sub
                frmODASPPaymentMethod.txtPaymentMethod = !PaymentMethod
                frmODASPPaymentMethod.txtPaymentMethodDescription = !PaymentMethodDescription
                
                If !AccountNoRequired = "Y" Or !AccountNoRequired = "Yes" Then
                        frmODASPPaymentMethod.optRequireAccountDetails.Value = True
                        frmODASPPaymentMethod.optJournal.Value = False
                        frmODASPPaymentMethod.optCash.Value = False
                
                ElseIf !Cash = "Y" Or !Cash = "Yes" Then
                        frmODASPPaymentMethod.optRequireAccountDetails.Value = False
                        frmODASPPaymentMethod.optJournal.Value = False
                        frmODASPPaymentMethod.optCash.Value = True
                
                ElseIf !Journal = "Y" Or !Journal = "Yes" Then
                        frmODASPPaymentMethod.optRequireAccountDetails.Value = False
                        frmODASPPaymentMethod.optJournal.Value = True
                        frmODASPPaymentMethod.optCash.Value = False
               Else
                        frmODASPPaymentMethod.optRequireAccountDetails.Value = False
                        frmODASPPaymentMethod.optJournal.Value = False
                        frmODASPPaymentMethod.optCash.Value = False

                
                End If
    End With

rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdDelete_Click()
''On Error GoTo Myerr
    
    If rep = True Then Exit Sub
    If frmODASPPaymentMethod.txtPaymentMethod.Text = Empty Then
        MsgBox "There is no current record to delete", vbInformation, "Delete Information"
    ElseIf frmODASPPaymentMethod.txtPaymentMethodDescription.Text = "" Then
        MsgBox "There is no current record to delete", vbInformation, "Delete Information"
    Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        
        Set rsCONTROL = New ADODB.Recordset
        rsCONTROL.Open "SELECT* FROM ODASPPaymentMethod WHERE PaymentMethod = '" & Me.txtPaymentMethod & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
            If .EOF And .BOF Then Exit Sub
            .Delete adAffectCurrent
            .Requery
            clearALLRECORD
            disableALLRECORD
            enableButtons
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
    disableButtons
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdCancel_Click()
        rep = False
        clearALLRECORD
        disableALLRECORD
        enableButtons
        frmODASPPaymentMethod.ListView1.ListItems.Clear
End Sub

Private Sub cmdUpdate_Click()
    rep = True
    bsaveRECORD = False
    validateRECORD
    If bsaveRECORD = True Then
        saveRecord
        loadGRID
        disableALLRECORD
        enableButtons
    End If
    rep = False


End Sub
Public Sub loadGRID()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Payment Method", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2


                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "select PaymentMethod, PaymentMethodDescription from ODASPPaymentMethod "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PaymentMethod))
                        
                        If Not IsNull(rsLIST!PaymentMethodDescription) Then
                            MyList.SubItems(1) = CStr(rsLIST!PaymentMethodDescription)
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

Private Sub validateRECORD()
On Error GoTo err
    If frmODASPPaymentMethod.txtPaymentMethod.Text = "" Then
                MsgBox "Enter the Payment Method ", vbCritical + vbOKOnly
                frmODASPPaymentMethod.txtPaymentMethod.SetFocus

    ElseIf frmODASPPaymentMethod.txtPaymentMethodDescription.Text = "" Then
                MsgBox "Enter the description", vbCritical + vbOKOnly
                frmODASPPaymentMethod.txtPaymentMethodDescription.SetFocus

    Else: bsaveRECORD = True
    
    End If

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub saveRecord()
On Error GoTo err
            
            strSQL = "select * from ODASPPaymentMethod where PaymentMethod = '" & frmODASPPaymentMethod.txtPaymentMethod.Text & "';"
            Set rsSAVE = New ADODB.Recordset
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            With rsSAVE
            
                If .BOF Or .EOF Then
                        .AddNew
                        !PaymentMethod = frmODASPPaymentMethod.txtPaymentMethod
                        !dateprepared = Date
                        !Preparedby = CurrentUserName
                End If
                
                !PaymentMethodDescription = frmODASPPaymentMethod.txtPaymentMethodDescription
                
                If frmODASPPaymentMethod.optCash.Value = True Then
                        !RequireAccountDetails = "N"
                        !Cash = "Y"
                        !Journal = "N"
                ElseIf frmODASPPaymentMethod.optJournal.Value = True Then
                        !RequireAccountDetails = "N"
                        !Cash = "N"
                        !Journal = "Y"
                ElseIf frmODASPPaymentMethod.optRequireAccountDetails.Value = True Then
                        !RequireAccountDetails = "Y"
                        !Cash = "N"
                        !Journal = "N"
                End If
                bsaveRECORD = False

                .Update
                .Requery
             rep = False
            
            End With
        
Exit Sub

err:

    If err.Number = -2147217842 Or err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
        rsSAVE.CancelUpdate
        rsSAVE.Requery
        cmdCancel.SetFocus
    Else
        ErrorMessage
    End If
End Sub

Private Sub cmdSearch_Click()
        enableButtons
        disableALLRECORD
End Sub

Private Sub Form_Activate()
    loadGRID
    disableALLRECORD
    enableButtons
    loadRECORD
End Sub

Private Sub Form_Load()
        OpenConnection
End Sub
Private Sub Form_Unload(cancel As Integer)
On Error GoTo err
    If rep = True Then
        cancel = True
        MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
    Else
        cancel = False
    End If
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i, j As Double
    
    If Item.Checked = True Then
        j = Me.ListView1.ListItems.Count
        
        If j = 0 Then Exit Sub
        
        For i = 1 To j
            If Me.ListView1.ListItems(i) <> Item Then
               Me.ListView1.ListItems(i).Checked = False
            End If
        Next i
                
    End If

        frmODASPPaymentMethod.txtPaymentMethod.Text = Item.Text
        loadRECORD

End Sub

