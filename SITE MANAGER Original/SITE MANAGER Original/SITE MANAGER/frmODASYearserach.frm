VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASYearserach 
   Caption         =   "Form1"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtEndDate 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtStartDate 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPickerStartDate 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Format          =   55377921
      CurrentDate     =   38298
   End
   Begin MSComCtl2.DTPicker DTPickerEndDate 
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Format          =   55377921
      CurrentDate     =   38298
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "End Date"
      Height          =   195
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Start Date"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmODASYearserach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DTPickerEndDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub DTPickerEndDate_CloseUp()
Me.txtEndDate.Text = Me.DTPickerEndDate
End Sub

Private Sub DTPickerStartDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub DTPickerStartDate_CloseUp()
Me.txtStartDate.Text = Me.DTPickerStartDate
End Sub
