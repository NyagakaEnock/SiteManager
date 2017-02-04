VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmALISRCommissionPeriod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commission Processing"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Reconstruct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   7
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   7215
      Begin VB.OptionButton Option1 
         Caption         =   "Premium Tax"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   2535
      End
      Begin VB.OptionButton optCommissionListing 
         Caption         =   "Commission Listing"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optCommissionStatement 
         Caption         =   "Commission Statement"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.ComboBox cboCurrentPeriod 
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
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   3375
   End
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
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Period "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   660
      Width           =   1455
   End
End
Attribute VB_Name = "frmALISRCommissionPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCOMM As clsCalculateCommission

Private Sub cmdPrint_Click()
'On Error GoTo err
    If frmALISRCommissionPeriod.optCommissionStatement = True Then
            Load frmALISRCommStatement
            frmALISRCommStatement.Show 1, ALISENTPMAIN
    ElseIf frmALISRCommissionPeriod.optCommissionListing = True Then
            Load frmALISRCommListing
            frmALISRCommListing.Show 1, ALISENTPMAIN
    End If
    
    
Exit Sub

err:
ErrorMessage
End Sub

Private Sub Command1_Click()
'On Error GoTo err
    If frmALISRCommissionPeriod.cboCurrentPeriod.Text > "" And CurrentUserName = "Administrator" Then

        Set rsCOMM = New clsCalculateCommission
        rsCOMM.CalculateBatchCommission
        Set rsCOMM = Nothing
    End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Load()
OpenConnection
LoadPeriod
frmALISRCommissionPeriod.ProgressBar1.Visible = False
End Sub

Private Sub LoadPeriod()
'On Error GoTo err
        Dim rsDEFAULT As ADODB.Recordset, strDEFAULT As String
        Set rsDEFAULT = New ADODB.Recordset
       
        strDEFAULT = "SELECT * FROM ALISPDefaults ; "
        rsDEFAULT.Open strDEFAULT, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsDEFAULT
            If .BOF Or .EOF Then Exit Sub
            
            frmALISRLoanPeriod.txtCurrentPeriod.Text = !CurrentPeriod
            
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboCurrentPeriod_GotFocus()

        Dim rsPERIOD As ADODB.Recordset, strPERIOD As String
        Set rsPERIOD = New ADODB.Recordset

        strPERIOD = "SELECT * FROM ALISPPeriod ; "
        rsPERIOD.Open strPERIOD, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        frmALISRCommissionPeriod.cboCurrentPeriod.Clear
        With rsPERIOD
            Do Until .EOF
                Screen.ActiveForm.cboCurrentPeriod.AddItem !AccountingPeriod
                .MoveNext
            Loop
        End With

rsPERIOD.Close
strPERIOD = ""

Exit Sub

err:
        ErrorMessage
End Sub

Private Sub Option2_Click()

End Sub
