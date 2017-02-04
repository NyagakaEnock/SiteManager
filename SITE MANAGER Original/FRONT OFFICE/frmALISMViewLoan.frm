VERSION 5.00
Begin VB.Form frmALISMViewLoan 
   Caption         =   "View Loan Details"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   8775
         Begin VB.TextBox txtReferenceNo 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7440
            TabIndex        =   45
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox cboLoanNo 
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
            Left            =   120
            TabIndex        =   44
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtApplicationNo 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6240
            TabIndex        =   43
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtNames 
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
            Left            =   3000
            TabIndex        =   42
            Top             =   480
            Width           =   3255
         End
         Begin VB.TextBox txtPolicyNo 
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
            Left            =   1440
            TabIndex        =   41
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Reference No"
            Height          =   195
            Left            =   7440
            TabIndex        =   50
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Names"
            Height          =   195
            Left            =   4200
            TabIndex        =   49
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Policy No"
            Height          =   240
            Left            =   1680
            TabIndex        =   48
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Loan No"
            Height          =   240
            Left            =   480
            TabIndex        =   47
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Application No"
            Height          =   240
            Left            =   6240
            TabIndex        =   46
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Present Details"
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   4080
         Width           =   8775
         Begin VB.TextBox txtInterestReceived 
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
            Left            =   840
            TabIndex        =   36
            Top             =   217
            Width           =   1455
         End
         Begin VB.TextBox txtTotalReceived 
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
            Left            =   5400
            TabIndex        =   35
            Top             =   217
            Width           =   1935
         End
         Begin VB.TextBox txtPrincipalReceived 
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
            Left            =   3120
            TabIndex        =   34
            Top             =   217
            Width           =   1695
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Interest"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   300
            Width           =   525
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   195
            Left            =   4920
            TabIndex        =   38
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Principal"
            Height          =   195
            Left            =   2400
            TabIndex        =   37
            Top             =   300
            Width           =   600
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cumulative Details"
         Height          =   735
         Left            =   120
         TabIndex        =   26
         Top             =   4800
         Width           =   8775
         Begin VB.TextBox txtPrincipalReceivedTodate 
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
            Left            =   3120
            TabIndex        =   29
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtTotalReceivedTodate 
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
            Left            =   5400
            TabIndex        =   28
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtInterestReceivedTodate 
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
            Left            =   840
            TabIndex        =   27
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Principal"
            Height          =   195
            Left            =   2400
            TabIndex        =   32
            Top             =   323
            Width           =   600
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   195
            Left            =   4920
            TabIndex        =   31
            Top             =   323
            Width           =   360
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Interest"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   323
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   8775
         Begin VB.TextBox txtPrincipalAmount 
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
            Left            =   6240
            TabIndex        =   13
            Top             =   645
            Width           =   1815
         End
         Begin VB.TextBox txtLoanAmount 
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
            Left            =   1680
            TabIndex        =   12
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtRepaymentPeriod 
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
            Left            =   6240
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtInterestDue 
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
            Left            =   6240
            TabIndex        =   10
            Top             =   1050
            Width           =   1815
         End
         Begin VB.TextBox txtInterestRate 
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
            Left            =   1680
            TabIndex        =   9
            Top             =   1050
            Width           =   2175
         End
         Begin VB.TextBox txtRepaymentAmount 
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
            Left            =   6240
            TabIndex        =   8
            Top             =   1455
            Width           =   1815
         End
         Begin VB.TextBox txtCommencementDate 
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
            Left            =   1680
            TabIndex        =   7
            Top             =   645
            Width           =   2175
         End
         Begin VB.TextBox txtStatus 
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
            Left            =   1680
            TabIndex        =   6
            Top             =   1860
            Width           =   2175
         End
         Begin VB.TextBox txtDueDate 
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
            Left            =   6240
            TabIndex        =   5
            Top             =   1860
            Width           =   1815
         End
         Begin VB.TextBox txtCurrentBalance 
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
            Left            =   1680
            TabIndex        =   4
            Top             =   1455
            Width           =   2175
         End
         Begin VB.TextBox txtDateOflastPayment 
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
            Left            =   1680
            TabIndex        =   3
            Top             =   2280
            Width           =   2175
         End
         Begin VB.TextBox txtCompleteDate 
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
            Left            =   6240
            TabIndex        =   2
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Principal  Amount"
            Height          =   195
            Left            =   4320
            TabIndex        =   25
            Top             =   728
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Loan Amount "
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   323
            Width           =   990
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Repayment Period (Mts.) "
            Height          =   195
            Left            =   4320
            TabIndex        =   23
            Top             =   323
            Width           =   1785
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Interest Due"
            Height          =   195
            Left            =   4320
            TabIndex        =   22
            Top             =   1133
            Width           =   870
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Interest (%)"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1133
            Width           =   780
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Repayment Amount"
            Height          =   195
            Left            =   4320
            TabIndex        =   20
            Top             =   1538
            Width           =   1395
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Commencement Date"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   728
            Width           =   1530
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   1943
            Width           =   450
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Due Date"
            Height          =   195
            Left            =   4320
            TabIndex        =   17
            Top             =   1943
            Width           =   690
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Current Balance"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1538
            Width           =   1140
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Last Payment Date"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   2340
            Width           =   1350
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Complete Date"
            Height          =   195
            Left            =   4320
            TabIndex        =   14
            Top             =   2370
            Width           =   1050
         End
      End
   End
End
Attribute VB_Name = "frmALISMViewLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVIEW As clsALISView


Private Sub Form_Activate()
    Set rsVIEW = New clsALISView
    rsVIEW.loadLOAN
    Set rsVIEW = Nothing
End Sub


