VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmALISPLoanType 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Loan Type"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   11250
   Icon            =   "frmALISPLoanType.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   11250
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame7 
         Height          =   1695
         Left            =   120
         TabIndex        =   47
         Top             =   4440
         Width           =   11055
         Begin MSDataGridLib.DataGrid BeneficiaryGrid 
            Height          =   1335
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   2355
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   8655
         Begin VB.TextBox txtMinimumRepaymentAmount 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   6120
            TabIndex        =   13
            Top             =   885
            Width           =   2055
         End
         Begin VB.TextBox txtMaximumRepaymentAmount 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   6120
            TabIndex        =   10
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtMinimumLoanPeriod 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   885
            Width           =   2055
         End
         Begin VB.TextBox txtMinimumLoanAmount 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   3960
            TabIndex        =   12
            Top             =   885
            Width           =   2055
         End
         Begin VB.TextBox txtMaximumLoanAmount 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   3960
            TabIndex        =   9
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtMaximumLoanPeriod 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   "Repayment Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            TabIndex        =   51
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Minimum"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   945
            Width           =   1095
         End
         Begin VB.Label lblMaximumRepaymentPeriod 
            Caption         =   "Loan Period"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   45
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lblContactTelephoneNo 
            Caption         =   "Maximum"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label lblTown 
            Caption         =   "Loan Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   43
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   8655
         Begin VB.ComboBox cboDebitAccount 
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
            Left            =   3480
            TabIndex        =   62
            Top             =   990
            Width           =   2430
         End
         Begin VB.ComboBox cboCreditAccount 
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
            Left            =   6960
            TabIndex        =   60
            Top             =   975
            Width           =   1590
         End
         Begin VB.ComboBox cboProductCode 
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
            Left            =   960
            TabIndex        =   58
            Top             =   990
            Width           =   1350
         End
         Begin VB.TextBox txtSurrenderPeriod 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   5160
            MaxLength       =   3
            TabIndex        =   3
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtLoanType 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   960
            MaxLength       =   5
            TabIndex        =   1
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtMonthlyInterestRate 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   3480
            MaxLength       =   3
            TabIndex        =   7
            Top             =   615
            Width           =   855
         End
         Begin VB.ComboBox cboStatus 
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
            Left            =   6960
            TabIndex        =   5
            Top             =   255
            Width           =   1590
         End
         Begin VB.TextBox txtLoanDescription 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   3480
            TabIndex        =   2
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtLoanRate 
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   960
            MaxLength       =   3
            TabIndex        =   4
            Top             =   615
            Width           =   1095
         End
         Begin VB.ComboBox cboFormula 
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
            Left            =   6960
            TabIndex        =   6
            Top             =   615
            Width           =   1590
         End
         Begin VB.Label Label12 
            Caption         =   "Debit AC"
            Height          =   255
            Left            =   2520
            TabIndex        =   63
            Top             =   1050
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Credit AC"
            Height          =   255
            Left            =   6000
            TabIndex        =   61
            Top             =   1035
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Product"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1050
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Surr Period"
            Height          =   255
            Left            =   4320
            TabIndex        =   57
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "%"
            Height          =   255
            Left            =   2040
            TabIndex        =   56
            Top             =   675
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Interest Rate"
            Height          =   255
            Left            =   2520
            TabIndex        =   49
            Top             =   675
            Width           =   975
         End
         Begin VB.Label lbLLoanRate 
            Caption         =   "Loan Rate"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   675
            Width           =   1215
         End
         Begin VB.Label lblTitle 
            Caption         =   "Formula"
            Height          =   255
            Left            =   6000
            TabIndex        =   40
            Top             =   668
            Width           =   615
         End
         Begin VB.Label lblSurname 
            Caption         =   "Description"
            Height          =   255
            Left            =   2520
            TabIndex        =   39
            Top             =   330
            Width           =   975
         End
         Begin VB.Label lblLoanType 
            AutoSize        =   -1  'True
            Caption         =   "Loan Type"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   330
            Width           =   765
         End
         Begin VB.Label lblFormula 
            Caption         =   "Status"
            Height          =   255
            Left            =   6000
            TabIndex        =   37
            Top             =   315
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   8880
         TabIndex        =   30
         Top             =   0
         Width           =   2295
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
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
            TabIndex        =   35
            Top             =   3315
            Width           =   2055
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
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
            TabIndex        =   34
            Top             =   2700
            Width           =   2055
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
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
            TabIndex        =   33
            Top             =   2085
            Width           =   2055
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
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
            TabIndex        =   32
            Top             =   1470
            Width           =   2055
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Update"
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
            TabIndex        =   22
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton cmdAddNew 
            Caption         =   "&Add New"
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
            TabIndex        =   31
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   25
         Top             =   6120
         Width           =   11055
         Begin VB.CommandButton cmdFirstCode 
            Height          =   375
            Index           =   0
            Left            =   2160
            Picture         =   "frmALISPLoanType.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdFirstCode 
            Height          =   375
            Index           =   1
            Left            =   3840
            Picture         =   "frmALISPLoanType.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdFirstCode 
            Height          =   375
            Index           =   2
            Left            =   5400
            Picture         =   "frmALISPLoanType.frx":0CC6
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdFirstCode 
            Height          =   375
            Index           =   3
            Left            =   6840
            Picture         =   "frmALISPLoanType.frx":1108
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1575
         Left            =   120
         TabIndex        =   24
         Top             =   2880
         Width           =   8655
         Begin VB.TextBox txtServiceChargeMinimumAmount 
            BackColor       =   &H00FFC0C0&
            Height          =   405
            Left            =   5160
            TabIndex        =   16
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtServiceChargeMaximumAmount 
            BackColor       =   &H00FFC0C0&
            Height          =   405
            Left            =   6840
            TabIndex        =   17
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtStampDutyMaximumAmount 
            BackColor       =   &H00FFC0C0&
            Height          =   405
            Left            =   6840
            TabIndex        =   21
            Top             =   1005
            Width           =   1575
         End
         Begin VB.ComboBox cboServiceChargeType 
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
            ItemData        =   "frmALISPLoanType.frx":154A
            Left            =   1680
            List            =   "frmALISPLoanType.frx":154C
            TabIndex        =   14
            Top             =   615
            Width           =   1575
         End
         Begin VB.TextBox txtServiceChargePercent 
            BackColor       =   &H00FFC0C0&
            Height          =   405
            Left            =   3360
            TabIndex        =   15
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtStampDutyMinimumAmount 
            BackColor       =   &H00FFC0C0&
            Height          =   405
            Left            =   5160
            TabIndex        =   20
            Top             =   1005
            Width           =   1575
         End
         Begin VB.ComboBox cboStampDutyType 
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
            TabIndex        =   18
            Top             =   975
            Width           =   1575
         End
         Begin VB.TextBox txtStampDutyPercent 
            BackColor       =   &H00FFC0C0&
            Height          =   405
            Left            =   3360
            TabIndex        =   19
            Top             =   1005
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   55
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Percent/Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   54
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Minimum Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   53
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Maximum Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6840
            TabIndex        =   52
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblServiceChargeType 
            Caption         =   "Service Charge"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lblStampDuty 
            Caption         =   "Stamp Duty"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1050
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmALISPLoanType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsLOANTYPE As cLoanTypes

Private Sub cboFormula_GotFocus()
On Error GoTo err:

        rsLOANTYPE.formulaGOTFOCUS

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboFormula_keypress(KeyAscii As Integer)

On Error GoTo err:
        KeyAscii = 0
Exit Sub

err:
        ErrorMessage
End Sub


Private Sub cboFormula_LostFocus()
On Error GoTo err:

        rsLOANTYPE.formulaLOSTFOCUS

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboProductCode_GotFocus()
On Error GoTo err:
    rsLOANTYPE.ProductCodeGotFocus
    
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub cboProductCode_KeyPress(KeyAscii As Integer)
On Error GoTo err:
        KeyAscii = 0
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboProductCode_LostFocus()
On Error GoTo err:
     rsLOANTYPE.ProductCodeLostFocus
Exit Sub

err:
    ErrorMessage
    
End Sub

Private Sub cboServiceChargeType_gotfocus()
On Error GoTo err:

        rsLOANTYPE.servicechargeGOTFOCUS

Exit Sub

err:
    ErrorMessage

End Sub
Private Sub cboservicechargetype_keypress(KeyAscii As Integer)
On Error GoTo err:

    KeyAscii = 0

Exit Sub

err:
    ErrorMessage
End Sub



Private Sub cboServiceChargeType_Lostfocus()
On Error GoTo err:

        rsLOANTYPE.servicechargeLOSTFOCUS

Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cboStampDutyType_GotFocus()
On Error GoTo err:

        rsLOANTYPE.stampdutyGOTFOCUS

Exit Sub

err:
    ErrorMessage

End Sub
Private Sub cbostampdutyType_keypress(KeyAscii As Integer)
On Error GoTo err:
        KeyAscii = 0
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cboStampDutyType_LostFocus()
On Error GoTo err:

        rsLOANTYPE.stampdutyLOSTFOCUS

Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cboStatus_gotFocus()
On Error GoTo err:

        rsLOANTYPE.statusGOTFOCUS

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cbostatus_keypress(KeyAscii As Integer)
On Error GoTo err:

        KeyAscii = 0
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cboStatus_LostFocus()
On Error GoTo err:

        rsLOANTYPE.statusLOSTFOCUS

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo err
    rsLOANTYPE.clearRECORD
    rsLOANTYPE.enableRECORD
    frmALISPLoanType.SetFocus
    rsLOANTYPE.DisableCommandButtons
    
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdCancel_Click()
On Error GoTo err
        rsLOANTYPE.cancelRECORD
         Exit Sub

err:
        ErrorMessage
End Sub

Private Sub cmdDelete_Click()
On Error GoTo err
        
        rsLOANTYPE.validateRECORD
        
        If rsLOANTYPE.BSave = True Then
            rsLOANTYPE.DeleteRecord
        End If

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub cmdFirstCode_Click(Index As Integer)
On Error GoTo err
    rsLOANTYPE.browseRECORD (Index)
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdSearch_Click()
On Error GoTo err:

        rsLOANTYPE.SearchRECORD

Exit Sub

err:
        ErrorMessage
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo err:

        rsLOANTYPE.validateRECORD
        
        If rsLOANTYPE.BSave = True Then
                rsLOANTYPE.SaveRecord
                rsLOANTYPE.BSave = False
        End If
        
        rsLOANTYPE.disableRECORD
        rsLOANTYPE.EnableCommandButtons
        
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub Form_Load()
'On Error GoTo err
  
        'create the instance of the data source class
        Set rsLOANTYPE = New cLoanTypes
        Call rsLOANTYPE.disableRECORD
        Call rsLOANTYPE.LoadGrid

Exit Sub
err:

ErrorMessage
End Sub

