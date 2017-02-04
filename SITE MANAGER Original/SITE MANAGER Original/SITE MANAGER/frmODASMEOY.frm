VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASMEOY 
   Caption         =   "End of Year Processing"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   LinkTopic       =   "Form2"
   ScaleHeight     =   5595
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   4920
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdeoy 
         BackColor       =   &H00C0C000&
         Caption         =   "Update Council Rates"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frmODASMEOY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEOY As clsODASCouncilRates
Private Sub cmdeoy_Click()
        frmODASMEOY.cmdeoy.Enabled = False
        Set rsEOY = New clsODASCouncilRates
        rsEOY.loadSTARTDATE_batch
        Set rsEOY = Nothing
End Sub

