VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEmployeesPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EMPLOYEES PERSONAL AND CONTACT INFORMATION"
   ClientHeight    =   5655
   ClientLeft      =   2850
   ClientTop       =   1860
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9150
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
         Top             =   0
         Width           =   2295
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "PERSONAL DETAILS"
      TabPicture(0)   =   "frmEmployeesPersonal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "CONTACT INFORMATION"
      TabPicture(1)   =   "frmEmployeesPersonal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   4575
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   8895
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   6480
            TabIndex        =   55
            Top             =   960
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55902209
            CurrentDate     =   37796
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   3240
            TabIndex        =   54
            Top             =   960
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55902209
            CurrentDate     =   37796
         End
         Begin VB.TextBox txtDepartment 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   3600
            TabIndex        =   37
            Top             =   3260
            Width           =   3135
         End
         Begin VB.ComboBox cboOfficialTitle 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1200
            TabIndex        =   34
            Top             =   3720
            Width           =   3615
         End
         Begin VB.ComboBox cboDepartment 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1200
            TabIndex        =   33
            Top             =   3260
            Width           =   2295
         End
         Begin VB.ComboBox cboGrade 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   4800
            TabIndex        =   30
            Top             =   2784
            Width           =   1935
         End
         Begin VB.ComboBox cboEmployType 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1200
            TabIndex        =   29
            Top             =   2784
            Width           =   2295
         End
         Begin VB.ComboBox cboGender 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1200
            TabIndex        =   26
            Top             =   1420
            Width           =   2295
         End
         Begin VB.ComboBox cboMaritalStatus 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   4800
            TabIndex        =   25
            Top             =   1425
            Width           =   1935
         End
         Begin VB.TextBox txtPINNumber 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   4800
            TabIndex        =   22
            Top             =   2340
            Width           =   1935
         End
         Begin VB.TextBox txtPassportNo 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1200
            TabIndex        =   21
            Top             =   2340
            Width           =   2295
         End
         Begin VB.TextBox txtNatIDNo 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   4800
            TabIndex        =   18
            Top             =   1880
            Width           =   1935
         End
         Begin VB.ComboBox cboNationality 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1200
            TabIndex        =   17
            Top             =   1880
            Width           =   2295
         End
         Begin VB.TextBox txtDateHired 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   4800
            TabIndex        =   14
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtDateofBirth 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1200
            TabIndex        =   13
            Top             =   960
            Width           =   2055
         End
         Begin MSComDlg.CommonDialog dlgPHOTO 
            Left            =   7320
            Top             =   2160
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "EMPLOYEE'S PHOTO"
            Filter          =   "Pictures (*.bmp;*.ico;*.jpg)|*.bmp;*.ico;*.jpg"
            FontBold        =   -1  'True
         End
         Begin VB.CommandButton cmdPHOTO 
            BackColor       =   &H80000009&
            Caption         =   "Load &PHOTO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3255
            Width           =   1935
         End
         Begin VB.ComboBox cboTitle 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   7080
            TabIndex        =   10
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton cmdNext 
            BackColor       =   &H80000001&
            Caption         =   "Next >>"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3960
            Width           =   1935
         End
         Begin VB.CheckBox chkStaffID 
            Caption         =   "Auto Staff ID No."
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
            TabIndex        =   8
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtOtherNames 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   4320
            TabIndex        =   5
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txtSurname 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   2160
            TabIndex        =   4
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtStaffIDNo 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label15 
            Caption         =   "Official Title"
            Height          =   315
            Left            =   120
            TabIndex        =   36
            Top             =   3720
            Width           =   1095
         End
         Begin VB.Label Label14 
            Caption         =   "Department"
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   3255
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Employ. Type"
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   2790
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Grade"
            Height          =   315
            Left            =   3600
            TabIndex        =   31
            Top             =   2790
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Marital Status"
            Height          =   315
            Left            =   3600
            TabIndex        =   28
            Top             =   1425
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Gender"
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   1425
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "P.I.N. Number"
            Height          =   315
            Left            =   3600
            TabIndex        =   24
            Top             =   2340
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Passport No"
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   2340
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "National ID No."
            Height          =   315
            Left            =   3600
            TabIndex        =   20
            Top             =   1875
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Nationality"
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   1875
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Date Hired"
            Height          =   315
            Left            =   3600
            TabIndex        =   16
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Date of Birth"
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Title of Courtesy"
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
            Left            =   7080
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
         Begin VB.Image imgPHOTO 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   2295
            Left            =   6840
            Stretch         =   -1  'True
            Top             =   960
            Width           =   1935
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   120
            X2              =   8760
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label2 
            Caption         =   "Surname"
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
            Left            =   2160
            TabIndex        =   7
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Other Names"
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
            Left            =   4320
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   8895
         Begin VB.CommandButton cmdFINISH 
            BackColor       =   &H80000001&
            Caption         =   "<< FINISH >>"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7095
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   3960
            Width           =   1695
         End
         Begin VB.CommandButton cmdBACK 
            BackColor       =   &H80000001&
            Caption         =   "<< BAC&K"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   3960
            Width           =   1335
         End
         Begin VB.TextBox txtEmail 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   2160
            TabIndex        =   56
            Top             =   3600
            Width           =   2415
         End
         Begin VB.TextBox txtMobileNo 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   2160
            TabIndex        =   47
            Top             =   3120
            Width           =   2415
         End
         Begin VB.TextBox txtTelephoneNo 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   2160
            TabIndex        =   46
            Top             =   2610
            Width           =   2415
         End
         Begin VB.ComboBox cboCountry 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   5400
            TabIndex        =   45
            Top             =   2100
            Width           =   2175
         End
         Begin VB.ComboBox cboTownCity 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   2160
            TabIndex        =   44
            Top             =   2100
            Width           =   2415
         End
         Begin VB.TextBox txtPostalAddress 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   2160
            TabIndex        =   43
            Top             =   1590
            Width           =   3615
         End
         Begin VB.TextBox txtPhysicalAddress 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   2160
            TabIndex        =   42
            Top             =   1080
            Width           =   5415
         End
         Begin VB.TextBox txtFullNames 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   3120
            TabIndex        =   39
            Top             =   480
            Width           =   4455
         End
         Begin VB.ComboBox cboStaffIDNo 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label24 
            Caption         =   "Email Address"
            Height          =   255
            Left            =   360
            TabIndex        =   57
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label Label23 
            Caption         =   "Country"
            Height          =   315
            Left            =   4680
            TabIndex        =   53
            Top             =   2100
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Mobile No"
            Height          =   315
            Left            =   360
            TabIndex        =   52
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label21 
            Caption         =   "Telephone No"
            Height          =   315
            Left            =   360
            TabIndex        =   51
            Top             =   2610
            Width           =   1575
         End
         Begin VB.Label Label20 
            Caption         =   "Town / City"
            Height          =   315
            Left            =   360
            TabIndex        =   50
            Top             =   2100
            Width           =   1575
         End
         Begin VB.Label Label19 
            Caption         =   "Postal Address"
            Height          =   315
            Left            =   360
            TabIndex        =   49
            Top             =   1590
            Width           =   1575
         End
         Begin VB.Label Label18 
            Caption         =   "Physical Address"
            Height          =   315
            Left            =   360
            TabIndex        =   48
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label17 
            Caption         =   "Staff ID Number"
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
            TabIndex        =   41
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label16 
            Caption         =   "Full Names"
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
            Left            =   3120
            TabIndex        =   40
            Top             =   240
            Width           =   1335
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   120
            X2              =   7560
            Y1              =   840
            Y2              =   840
         End
      End
   End
End
Attribute VB_Name = "frmEmployeesPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyEmpPersonal As clsEmpPersonal

Private Sub cboCountry_Click()
    Me.txtTelephoneNo.SetFocus
End Sub

Private Sub cboCountry_GotFocus()
    MyEmpPersonal.GetCountryCode
End Sub


Private Sub cboCountry_LostFocus()
    MyEmpPersonal.FindCountryCode
End Sub

Private Sub cboDepartment_Click()
    Me.cboOfficialTitle.SetFocus
End Sub

Private Sub cboDepartment_GotFocus()
    MyEmpPersonal.GetDepartments
End Sub

Private Sub cboDepartment_LostFocus()
    MyEmpPersonal.FindDeptCode
End Sub

Private Sub cboEmployType_Click()
    Me.cboGrade.SetFocus
End Sub

Private Sub cboEmployType_GotFocus()
    MyEmpPersonal.GetEmploymentTypes
End Sub

Private Sub cboEmployType_LostFocus()
    MyEmpPersonal.FindEmpTypeCode
End Sub

Private Sub cboGender_Click()
    Me.cboMaritalStatus.SetFocus
End Sub

Private Sub cboGrade_Click()
    Me.cboDepartment.SetFocus
End Sub

Private Sub cboGrade_GotFocus()
    MyEmpPersonal.GetJobGroups
End Sub

Private Sub cboMaritalStatus_Click()
    Me.cboNationality.SetFocus
End Sub

Private Sub cboNationality_Click()
    Me.txtNatIDNo.SetFocus
End Sub

Private Sub cboNationality_GotFocus()
    MyEmpPersonal.GetNationality
End Sub

Private Sub cboNationality_LostFocus()
    MyEmpPersonal.FindNationalityCode
End Sub

Private Sub cboOfficialTitle_Click()
    Me.cmdNext.SetFocus
End Sub

Private Sub cboOfficialTitle_GotFocus()
    MyEmpPersonal.GetDesignations
End Sub

Private Sub cboOfficialTitle_LostFocus()
    MyEmpPersonal.FindDesignationCode
End Sub

Private Sub cboStaffIDNo_Click()
    Me.txtPhysicalAddress.SetFocus
End Sub

Private Sub cboStaffIDNo_GotFocus()
If NewRecord Or beditRECORD Then
    MyEmpPersonal.GetEmploymentTypes
    If Me.cboStaffIDNo.Text = Empty Then
        Me.cboStaffIDNo.Text = Trim(Me.txtStaffIDNo.Text)
        Me.txtFullNames.Text = Trim(Me.txtCompanyName.Text) & " " & Trim(Me.txtOtherNames.Text)
        Me.txtPhysicalAddress.SetFocus
    Else
        Exit Sub
    End If
Else
    Exit Sub
End If
End Sub

Private Sub cboStaffIDNo_LostFocus()
    MyEmpPersonal.FindFullNames
End Sub

Private Sub cboTitle_Click()
    Me.txtDateofBirth.SetFocus
End Sub

Private Sub cboTitle_GotFocus()
    MyEmpPersonal.GetTitles
End Sub

Private Sub cboTownCity_Click()
    Me.cboCountry.SetFocus
End Sub

Private Sub cboTownCity_GotFocus()
    MyEmpPersonal.GetMainCity
End Sub

Private Sub cmdAddNew_Click()
'On Error GoTo err
    Select Case cmdAddNew.Caption
    Case "&New"
        MyEmpPersonal.ClearMyScreen
        MyEmpPersonal.AddNewRecord
    Case "SAVE &RECORD"
        MyEmpPersonal.SaveNewRecord
    Case Else
        Exit Sub
    End Select
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdBACK_Click()
'On Error GoTo err
    If NewRecord Or beditRECORD Then
        Me.SSTab1.Tab = 0
        Me.txtStaffIDNo.SetFocus
    Else
        Exit Sub
    End If
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdEditRecord_Click()
'On Error GoTo err
Select Case cmdEditRecord.Caption
    Case "E&dit"
        MyEmpPersonal.CheckEditRecord
    Case "SAVE &CHANGES"
        MyEmpPersonal.EditMyRecord
    Case Else
        Exit Sub
    End Select
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdFINISH_Click()
'On Error GoTo err
    If NewRecord = True And beditRECORD = False Then
        If MsgBox("Are you Satisfied with all Entries Made?", vbQuestion + vbYesNo + vbDefaultButton1, "Save New Record") = vbNo Then Exit Sub
        Call cmdAddNew_Click
    ElseIf beditRECORD = True And NewRecord = False Then
        If MsgBox("Are you Satisfied with all Changes Made on the Current Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Save Changes") = vbNo Then Exit Sub
        Call cmdEditRecord_Click
    Else
        Exit Sub
    End If
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdNext_Click()
'On Error GoTo err
If NewRecord Or beditRECORD Then
    Me.SSTab1.Tab = 1
    Me.cboStaffIDNo.Text = Me.txtStaffIDNo.Text
    Me.txtFullNames.Text = Trim(Me.txtCompanyName.Text) & " " & Trim(Me.txtOtherNames.Text)
    Me.txtPhysicalAddress.SetFocus
Else
    Exit Sub
End If
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdPHOTO_Click()
    MyEmpPersonal.ShowOpenDialog
End Sub

Private Sub cmdRefresh_Click()
    MyEmpPersonal.RefreshScreen
End Sub

Private Sub DTPicker1_CloseUp()
If Not NewRecord And Not beditRECORD Then Exit Sub
    If Me.DTPicker1.Value > Date Then
        MsgBox "Date of Birth Cannot be in the Future!", vbCritical + vbOKOnly, "Invalid Date"
        Me.DTPicker1.Value = Date
        Me.DTPicker1.Refresh
        Exit Sub
    Else
        Me.txtDateofBirth.Text = Me.DTPicker1.Value
        Me.txtDateHired.SetFocus
    End If
End Sub

Private Sub DTPicker2_CloseUp()
If Not NewRecord And Not beditRECORD Then Exit Sub
    If Me.DTPicker2.Value > Date Then
        MsgBox "Date of Hire Cannot be in the Future!", vbCritical + vbOKOnly, "Invalid Date"
        Me.DTPicker2.Value = Date
        Me.DTPicker2.Refresh
        Exit Sub
    Else
        Me.txtDateHired.Text = Me.DTPicker2.Value
        Me.cboGender.SetFocus
    End If
End Sub

Private Sub Form_Load()
'On Error GoTo err
    Call OpenConnection
    Set MyEmpPersonal = New clsEmpPersonal
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If NewRecord Or beditRECORD Then MsgBox "Data Entry or Edit in Progress! No Work was Done!", vbInformation + vbOKOnly, "Screen Unload": Cancel = 1
End Sub

Private Sub txtPostalAddress_GotFocus()
If NewRecord Or beditRECORD Then
    If Me.txtPostalAddress.Text = Empty Then
        Me.txtPostalAddress.Text = "P.O. BOX ": Me.txtPostalAddress.SetFocus
        Me.txtPostalAddress.SelStart = Len(Me.txtPostalAddress.Text) + 1: Me.txtPostalAddress.SelLength = 0
    Else
        Me.txtPostalAddress.SetFocus
    End If
Else
    Exit Sub
End If
End Sub
