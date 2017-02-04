VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmALISMProceeds 
   Caption         =   "Proceeds/Deductions"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5318
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
         BackColor       =   16777152
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
End
Attribute VB_Name = "frmALISMProceeds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPAYMENT As clsALISPaymentRequisition


Private Sub Form_Activate()
    Set rsPAYMENT = New clsALISPaymentRequisition
    rsPAYMENT.loadPaymentGRID
    Set rsPAYMENT = Nothing
    
End Sub

