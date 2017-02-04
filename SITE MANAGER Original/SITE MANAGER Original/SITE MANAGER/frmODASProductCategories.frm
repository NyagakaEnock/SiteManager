VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPProductCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERAL INVENTORY-Product Categories"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   1680
   ClientWidth     =   9360
   Icon            =   "frmODASProductCategories.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9360
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H80000000&
         Caption         =   "&PRINT"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   2175
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H80000000&
         Caption         =   "REFRE&SH"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4720
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Clear screen"
         Top             =   0
         Width           =   2535
      End
      Begin VB.CommandButton cmdAddNew 
         BackColor       =   &H80000000&
         Caption         =   "&NEW"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Add New"
         Top             =   0
         Width           =   2535
      End
      Begin VB.CommandButton cmdEditRecord 
         BackColor       =   &H80000000&
         Caption         =   "E&DIT"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2550
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Change existing record"
         Top             =   0
         Width           =   2220
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   9375
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   360
         TabIndex        =   10
         ToolTipText     =   "List of current Product categories"
         Top             =   1440
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtCategoryName 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5760
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtCategoryCode 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Caption         =   "List Of Current Product Categories"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   9015
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   9120
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Category Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Category Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmODASPProductCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyDrugCategories As clsODASProductInfo

Private Sub cmdAddNew_Click()
If Edit = False Then
    MyDrugCategories.AddNewCat
End If
End Sub

Private Sub cmdEditRecord_Click()
    If Save = False Then
        MyDrugCategories.EditRecordCat
    End If
End Sub


Private Sub cmdPrint_Click()
Load frmRPTProductCategories
frmRPTProductCategories.Show 1, Me
End Sub

Private Sub cmdRefresh_Click()
    MyDrugCategories.RefreshCat
    Found = False
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    showProductCategory
End Sub

Private Sub Form_Initialize()
    Set MyDrugCategories = New clsODASProductInfo
End Sub

Private Sub Form_Load()
Call OpenODBCConnection

'On Error Resume Next
    Me.Left = 2805
    Me.Top = 1425
End Sub

Private Sub Form_LostFocus()
    Set MyDrugCategories = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo err

    If Save = True Or Edit = True Then
        MsgBox "Please there is Work going on, Refresh to continue", vbOKCancel + vbCritical
        Cancel = 1
    Else
         Found = False
    End If
    
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub txtCategoryCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Set MyDrugCategories = New clsODASProductInfo
MyDrugCategories.SearchRecordCat
End If
End Sub






