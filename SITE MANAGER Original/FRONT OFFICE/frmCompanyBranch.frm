VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCompanyBranch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Branch Network"
   ClientHeight    =   4950
   ClientLeft      =   2850
   ClientTop       =   1860
   ClientWidth     =   8610
   Icon            =   "frmCompanyBranch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8610
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   6600
      TabIndex        =   9
      Top             =   3060
      Width           =   1935
   End
   Begin VB.TextBox txtFaxTelex 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   3960
      TabIndex        =   8
      Top             =   3060
      Width           =   1935
   End
   Begin VB.TextBox txtPhoneNo 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   3060
      Width           =   1695
   End
   Begin VB.ComboBox cboContactTitle 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   12
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox txtStaffIDNo 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   5520
      TabIndex        =   11
      Top             =   3570
      Width           =   3015
   End
   Begin VB.TextBox txtContactName 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Top             =   3570
      Width           =   3255
   End
   Begin VB.ComboBox cboTownCity 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   5520
      TabIndex        =   6
      Top             =   2580
      Width           =   3015
   End
   Begin VB.TextBox txtPostAdd 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   2580
      Width           =   3255
   End
   Begin VB.TextBox txtPhyAdd 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   2100
      Width           =   7335
   End
   Begin VB.TextBox txtBranchName 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   1620
      Width           =   4695
   End
   Begin VB.TextBox txtBranchCode 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   1620
      Width           =   1935
   End
   Begin VB.ComboBox cboCompanyCode 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtCompanyName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   5295
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H80000001&
         Caption         =   "Refres&h"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddNew 
         BackColor       =   &H80000001&
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Add New"
         Top             =   0
         Width           =   2295
      End
      Begin VB.CommandButton cmdEditRecord 
         BackColor       =   &H80000001&
         Caption         =   "E&dit"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Label Label13 
      Caption         =   "E-Mail"
      Height          =   255
      Left            =   6000
      TabIndex        =   29
      Top             =   3060
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Fax / Telex"
      Height          =   255
      Left            =   3000
      TabIndex        =   28
      Top             =   3060
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Phone No."
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Contact Title"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Staff ID NO"
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   3570
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Contact Name"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3570
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Town / City"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Postal Address"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2580
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Physical  Add."
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   1620
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Branch Code"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Name of Company"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Company Code"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   8520
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmCompanyBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyBranches As clsCompanyBranch

Private Sub cmdAddNew_Click()
On Error GoTo Err
    Select Case cmdAddNew.Caption
    Case "&New"
        MyBranches.ClearMyScreen
        MyBranches.GetCompanyInfo
        MyBranches.AddNewRecord
    Case "SAVE &RECORD"
        MyBranches.SaveNewRecord
    Case Else
        Exit Sub
    End Select
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cmdEditRecord_Click()
On Error GoTo Err
Select Case cmdEditRecord.Caption
    Case "E&dit"
        MyBranches.CheckEditRecord
    Case "SAVE &CHANGES"
        MyBranches.EditMyRecord
    Case Else
        Exit Sub
    End Select
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cmdRefresh_Click()
    MyBranches.RefreshScreen
End Sub

Private Sub Form_Load()
On Error GoTo Err
    Call OpenConnection
    Set MyBranches = New clsCompanyBranch
    
Exit Sub
Err:
    ErrorMessage
End Sub
