VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   8460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4530
      Left            =   -90
      TabIndex        =   0
      Top             =   -180
      Width           =   8520
      Begin VB.Label lblCopyright 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Copyright: AUGUST 2003"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5280
         TabIndex        =   2
         Top             =   3780
         Width           =   3135
      End
      Begin VB.Label lblCompany 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Company: Fourtune Technologies Ltd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5280
         TabIndex        =   1
         Top             =   4110
         Width           =   3135
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Version: 1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   345
         Left            =   6720
         TabIndex        =   3
         Top             =   2700
         Width           =   1620
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "for Microsoft Windows 98/Me/2000/XP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   480
         Left            =   2115
         TabIndex        =   4
         Top             =   2220
         Width           =   6225
      End
      Begin VB.Label lblProductName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "LIBRARY CENTRAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   765
         Left            =   2115
         TabIndex        =   6
         Top             =   1260
         Width           =   6360
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "THE GOLDENEYE INTEGRATED LIBRARY MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   840
         Left            =   2115
         TabIndex        =   5
         Top             =   225
         Width           =   6420
      End
      Begin VB.Image Image1 
         Height          =   4290
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2010
      End
      Begin VB.Image Image2 
         Height          =   1530
         Left            =   7080
         Picture         =   "frmSplash.frx":68EE
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1410
      End
      Begin VB.Image imgLogo 
         Height          =   4305
         Left            =   120
         Picture         =   "frmSplash.frx":D1D0
         Stretch         =   -1  'True
         Top             =   195
         Width           =   8655
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Call OpenConnection

End Sub
