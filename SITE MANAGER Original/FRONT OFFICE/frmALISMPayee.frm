VERSION 5.00
Begin VB.Form frmALISMPayee 
   Caption         =   "Payee Details"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Frame Frame5 
         Caption         =   "Payee Details"
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6015
         Begin VB.TextBox txtPayeeAddress 
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
            Left            =   360
            TabIndex        =   5
            Top             =   735
            Width           =   5295
         End
         Begin VB.ComboBox cboTownCode 
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
            Left            =   360
            TabIndex        =   4
            Top             =   1920
            Width           =   5295
         End
         Begin VB.TextBox txtPostalCode 
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
            Left            =   360
            TabIndex        =   3
            Top             =   1320
            Width           =   5295
         End
         Begin VB.TextBox txtPayeeDetails 
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
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   5295
         End
      End
   End
End
Attribute VB_Name = "frmALISMPayee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPAYMENT As clsALISPaymentRequisition

Private Sub Form_Activate()
    Set rsPAYMENT = New clsALISPaymentRequisition
    rsPAYMENT.loadPAYEE
    Set rsPAYMENT = Nothing
    
End Sub
