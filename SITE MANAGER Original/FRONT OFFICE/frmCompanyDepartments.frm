VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCompanyDepartments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPANY DEPARTMENTS"
   ClientHeight    =   5055
   ClientLeft      =   2850
   ClientTop       =   1860
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8895
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
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
         TabIndex        =   16
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
         TabIndex        =   15
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
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   8895
      Begin VB.TextBox txtDeptNotes 
         BackColor       =   &H00FFC0C0&
         Height          =   1335
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2400
         Width           =   5175
      End
      Begin VB.TextBox txtStaffID 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   5160
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtOfficialTitle 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   4
         Top             =   1860
         Width           =   5175
      End
      Begin VB.TextBox txtDeptName 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtDeptCode 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboDeptHead 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Comments / Notes"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Staff ID"
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   1320
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label5 
         Caption         =   "Department Head"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Department Name"
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
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Department Code"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Official Title"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1860
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCompanyDepartments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyDept As clsCompanyDepts

Private Sub cboDeptHead_Click()
If Not NewRecord And Not beditRECORD Then Exit Sub
    Me.txtDeptNotes.SetFocus
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo err
    Select Case cmdAddNew.Caption
    Case "&New"
        MyDept.ClearMyScreen
        MyDept.AddNewRecord
    Case "SAVE &RECORD"
        MyDept.SaveNewRecord
    Case Else
        Exit Sub
    End Select
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdEditRecord_Click()
On Error GoTo err
Select Case cmdEditRecord.Caption
    Case "E&dit"
        MyDept.CheckEditRecord
    Case "SAVE &CHANGES"
        MyDept.EditMyRecord
    Case Else
        Exit Sub
    End Select
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdRefresh_Click()
    MyDept.RefreshScreen
End Sub

Private Sub Form_Load()
On Error GoTo err
    Call OpenConnection
    Set MyDept = New clsCompanyDepts
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If NewRecord Or beditRECORD Then MsgBox "Data Entry or Edit in Progress! No Work was Done!", vbInformation + vbOKOnly, "Screen Unload": Cancel = 1
End Sub

