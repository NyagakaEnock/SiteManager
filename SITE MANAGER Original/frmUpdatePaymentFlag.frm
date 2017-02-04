VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmUpdatePaymentFlag 
   Caption         =   "Update Payment Flay"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C000&
         Caption         =   "Generate"
         Height          =   855
         Left            =   1800
         MaskColor       =   &H00C0C000&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmUpdatePaymentFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        updatePaymentFlag
End Sub
