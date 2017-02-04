VERSION 5.00
Begin VB.Form frmALISMReceiptCopy 
   Caption         =   "Print Receipts"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5190
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameSpecificClaim 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4935
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT RECEIPT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         Picture         =   "frmALISMReceiptCopy.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtReceiptAmount 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtReceiptNo 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Receipt Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Receipt No"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmALISMReceiptCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPrint_Click()
On Error GoTo err
    If frmALISMReceiptCopy.txtReceiptNo.Text > "" Then
            Load frmALISRReceiptCopy
            frmALISRReceiptCopy.Show 1, Me
    End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadRECORD()
On Error GoTo err

    Set rsCONTROL = New ADODB.Recordset
    strSQL = "Select * from ALISMReceiptNew Where ReceiptNo = '" & frmALISMReceiptCopy.txtReceiptNo.Text & "' and PaymentStatus = 'PAID';"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsCONTROL
            If .BOF Or .EOF Then Exit Sub
            frmALISMReceiptCopy.txtReceiptAmount.Text = !ReceiptAmount
    End With
    
strSQL = ""
rsCONTROL.Close

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
    loadRECORD
End Sub

