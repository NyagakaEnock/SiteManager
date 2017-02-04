VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPANY / ORGANISATIONAL INFORMATION"
   ClientHeight    =   5415
   ClientLeft      =   2850
   ClientTop       =   1860
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9255
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      _Version        =   393216
      BorderStyle     =   1
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
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   0
         Width           =   2415
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
         TabIndex        =   27
         ToolTipText     =   "Add New"
         Top             =   0
         Width           =   2415
      End
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
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   0
      TabIndex        =   12
      Top             =   360
      Width           =   9255
      Begin VB.TextBox txtCoyNHIFNo 
         BackColor       =   &H00FFC0C0&
         Height          =   360
         Left            =   1680
         TabIndex        =   11
         Top             =   4440
         Width           =   3015
      End
      Begin VB.TextBox txtCoyPinNo 
         BackColor       =   &H00FFC0C0&
         Height          =   360
         Left            =   1680
         TabIndex        =   10
         Top             =   3960
         Width           =   3015
      End
      Begin VB.TextBox txtCoyITNo 
         BackColor       =   &H00FFC0C0&
         Height          =   360
         Left            =   1680
         TabIndex        =   9
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox txtCompanyName 
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
         Left            =   2640
         TabIndex        =   1
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtAddress1 
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
         Left            =   1680
         TabIndex        =   2
         Top             =   1169
         Width           =   5535
      End
      Begin VB.TextBox txtAddress2 
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
         Left            =   1680
         TabIndex        =   3
         Top             =   1626
         Width           =   5535
      End
      Begin VB.TextBox txtCompanyNo 
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
         Width           =   2415
      End
      Begin VB.TextBox txtTelephone 
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
         Left            =   1680
         TabIndex        =   6
         Top             =   2540
         Width           =   2055
      End
      Begin VB.TextBox txtFaxTelex 
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
         Left            =   4680
         TabIndex        =   7
         Top             =   2540
         Width           =   2535
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   1680
         TabIndex        =   8
         Top             =   3000
         Width           =   3015
      End
      Begin VB.ComboBox cboTownCity 
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
         Left            =   1680
         TabIndex        =   4
         Top             =   2083
         Width           =   2055
      End
      Begin VB.ComboBox cboCountry 
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
         Left            =   4680
         TabIndex        =   5
         Top             =   2083
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "NHIF Number"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "PIN Number"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "I.T. Number"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Postal Address"
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   1626
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Physical Address"
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   1169
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name of Company / Institution"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   18
         Top             =   360
         Width           =   2340
      End
      Begin VB.Label Label10 
         Caption         =   "Fax/Telex:"
         Height          =   315
         Left            =   3840
         TabIndex        =   17
         Top             =   2540
         Width           =   825
      End
      Begin VB.Label Label11 
         Caption         =   "Telephone:"
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   2540
         Width           =   1305
      End
      Begin VB.Label Label12 
         Caption         =   "E- Mail:"
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   1305
      End
      Begin VB.Label Label9 
         Caption         =   "Country:"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   2083
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Town/City:"
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   2083
         Width           =   1305
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   7200
         Y1              =   1050
         Y2              =   1050
      End
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyCompany As clsCompany

Private Sub cmdAddNew_Click()
On Error GoTo err
    Select Case cmdAddNew.Caption
    Case "&New"
        MyCompany.ClearMyScreen
        MyCompany.AddNewRecord
    Case "SAVE &RECORD"
        MyCompany.SaveNewRecord
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
        MyCompany.CheckEditRecord
    Case "SAVE &CHANGES"
        MyCompany.EditMyRecord
    Case Else
        Exit Sub
    End Select
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdRefresh_Click()
    MyCompany.RefreshScreen
End Sub

Private Sub Form_Load()
On Error GoTo err
    Call OpenConnection
    Set MyCompany = New clsCompany
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboTownCity_GotFocus()
    MyCompany.GetTownCity
End Sub

Private Sub cboCountry_GotFocus()
    MyCompany.GetCountryCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If NewRecord Or beditRECORD Then MsgBox "Data Entry or Edit in Progress! No Work was Done!", vbInformation + vbOKOnly, "Screen Unload": Cancel = 1
End Sub
