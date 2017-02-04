VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmALISMCheckOff 
   Caption         =   "Stop Order Processing"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   120
         TabIndex        =   43
         Top             =   4440
         Width           =   9015
         Begin MSComctlLib.ListView ListView3 
            Height          =   2415
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   4260
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
      Begin VB.Frame Frame3 
         Height          =   2895
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   7695
         Begin VB.TextBox txtItemAccountCode 
            Alignment       =   1  'Right Justify
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
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   27
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox txtDeductionCode 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   26
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtPersonalNumber 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   25
            Top             =   1380
            Width           =   1935
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
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
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   24
            Top             =   1380
            Width           =   1815
         End
         Begin VB.TextBox txtRecurrentAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   23
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox txtAmountOff 
            Alignment       =   1  'Right Justify
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
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   22
            Top             =   540
            Width           =   1815
         End
         Begin VB.TextBox txtReference 
            Alignment       =   1  'Right Justify
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
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   21
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtCheckDigit 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   20
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox txtAmountOn 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   19
            Top             =   540
            Width           =   1935
         End
         Begin VB.TextBox txtCode 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   18
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtIdentityDetails 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   63
            TabIndex        =   17
            Top             =   2280
            Width           =   5655
         End
         Begin VB.Label Label16 
            Caption         =   "Item Account Code"
            Height          =   255
            Left            =   4080
            TabIndex        =   38
            Top             =   1860
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Deduction Code"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   1860
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Personal Number"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Balance "
            Height          =   255
            Left            =   4080
            TabIndex        =   35
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "Recurrent Amount"
            Height          =   255
            Left            =   4080
            TabIndex        =   34
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Amount Off"
            Height          =   255
            Left            =   4080
            TabIndex        =   33
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "Reference No"
            Height          =   255
            Left            =   4080
            TabIndex        =   32
            Top             =   1020
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Check Digit"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Amount On"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Code"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1020
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Identity Details"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   2280
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   7695
         Begin VB.TextBox txtDepartmentDescription 
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1020
            Width           =   5655
         End
         Begin VB.TextBox txtEmployerDescription 
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   600
            Width           =   5655
         End
         Begin VB.TextBox txtMonth 
            BackColor       =   &H00FFC0C0&
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
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   13
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtReferenceDate 
            BackColor       =   &H00FFC0C0&
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
            Left            =   5520
            TabIndex        =   12
            Top             =   240
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DTPickerReferenceDate 
            Height          =   375
            Left            =   7320
            TabIndex        =   11
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55902209
            CurrentDate     =   38149
         End
         Begin VB.Label Label6 
            Caption         =   "Department"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Employer"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   660
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Month"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Reference Date"
            Height          =   255
            Left            =   4200
            TabIndex        =   14
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.Frame fraCButtons 
         Height          =   4455
         Index           =   6
         Left            =   7920
         TabIndex        =   1
         Top             =   120
         Width           =   1215
         Begin VB.CommandButton cmdCancel 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "frmALISMCheckOff.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2810
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "frmALISMCheckOff.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2296
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdate 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "frmALISMCheckOff.frx":0204
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   754
            Width           =   975
         End
         Begin VB.CommandButton cmdEdit 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "frmALISMCheckOff.frx":0306
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1782
            Width           =   975
         End
         Begin VB.CommandButton cmdAddNew 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "frmALISMCheckOff.frx":0408
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdSearch 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "frmALISMCheckOff.frx":050A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1268
            Width           =   975
         End
         Begin VB.CommandButton cmdPrint 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "frmALISMCheckOff.frx":060C
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   3324
            Width           =   975
         End
         Begin VB.CommandButton cmdPendingDepts 
            Caption         =   "Pending Depts"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   3840
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmALISMCheckOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDepartmentCode_GotFocus()
    
    If frmALISMDeductionForm.cboAccountNo.Text <= "" Then Exit Sub
    
    strSQL = ""
    strSQL = "SELECT * FROM ALISPDepartment where AccountNo = '" & frmALISMDeductionForm.cboAccountNo.Text & "' and MonthPrepared <> '" & frmALISMDeductionForm.txtMonth.Text & "' OR MONTHPREPARED is null;"
    selectDepartmentCodeGotFocus
End Sub

Private Sub cboDepartmentCode_LostFocus()
    selectDepartmentCodeLostFocus
End Sub

Private Sub cboAccountNo_GotFocus()
    If frmALISMDeductionForm.txtMonth.Text <= "" Then Exit Sub
    strSQL = ""
    strSQL = "SELECT * FROM ODASPAccount WHERE StatusCode = 'A' and EmployerType = 'TSC';"
    selectAccountNoGotFocus
    
    Set rsLOADGRID = New clsALISGRID
    rsLOADGRID.loadDEPARTMENTGRID
    rsLOADGRID.loadPENDINGDEPTGRID
    rsLOADGRID.loadPREPAREDDEPTGRID
    Set rsLOADGRID = Nothing

End Sub

Private Sub cboAccountNo_LostFocus()
    selectAccountNoLostFocus
    
    Set rsLOADGRID = New clsALISGRID
    rsLOADGRID.loadDEPARTMENTGRID
    rsLOADGRID.loadPENDINGDEPTGRID
    rsLOADGRID.loadPREPAREDDEPTGRID
    Set rsLOADGRID = Nothing

End Sub

