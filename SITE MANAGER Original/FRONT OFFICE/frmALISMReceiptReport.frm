VERSION 5.00
Begin VB.Form frmALISMReceiptReport 
   Caption         =   "Receipts Reports"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5190
   LinkTopic       =   "Form2"
   ScaleHeight     =   2415
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameSpecificClaim 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4935
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT REPORT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1320
         Picture         =   "frmALISMReceiptReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "SetPeriod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cboAccountingPeriod 
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
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Accounting Period"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmALISMReceiptReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub cboAccountingPeriod_GotFocus()
On Error GoTo err
    
    Dim rsPAY As ADODB.Recordset, strPAY As String
    Set rsPAY = New Recordset
    
    strPAY = "SELECT * FROM ALISPPeriod;"
    rsPAY.Open strPAY, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboAccountingPeriod.Clear

    With rsPAY
            Do Until .EOF
                    Screen.ActiveForm.cboAccountingPeriod.AddItem !AccountingPeriod
                    .MoveNext
            Loop
    
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboAccountingPeriod_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdPrint_Click()
On Error GoTo err
    If frmALISMReceiptReport.cboAccountingPeriod.Text > "" Then
            Load frmReceipt
            frmReceipt.Show 1, Me
    End If
Exit Sub

err:
    ErrorMessage
End Sub

