VERSION 5.00
Begin VB.Form frmALISMViewPolicy 
   Caption         =   "View Policy Details"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame8 
         Height          =   1695
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   8415
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
            Height          =   375
            Left            =   1560
            TabIndex        =   41
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtLifeAssuredP 
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
            Height          =   375
            Left            =   4800
            TabIndex        =   40
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txtEmployerCodeP 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   39
            Top             =   660
            Width           =   2055
         End
         Begin VB.TextBox txtEmployerP 
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
            Height          =   375
            Left            =   4800
            TabIndex        =   38
            Top             =   660
            Width           =   3255
         End
         Begin VB.TextBox txtEmployeeNoP 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   37
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtReferenceNoP 
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
            Height          =   375
            Left            =   5520
            TabIndex        =   36
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label5 
            Caption         =   "Policy No"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Life Assured"
            Height          =   255
            Left            =   3840
            TabIndex        =   46
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "Employer Code"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label22 
            Caption         =   "Employer"
            Height          =   255
            Left            =   3840
            TabIndex        =   44
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Employee No"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1140
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Reference No "
            Height          =   255
            Left            =   3840
            TabIndex        =   42
            Top             =   1140
            Width           =   1335
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame8"
         Height          =   2175
         Left            =   120
         TabIndex        =   18
         Top             =   3960
         Width           =   8415
         Begin VB.TextBox txtReceivedTodateP 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   26
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtSuspenseAccountP 
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
            Height          =   375
            Left            =   5520
            TabIndex        =   25
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtDateOfLastPaymentP 
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
            Height          =   375
            Left            =   5520
            TabIndex        =   24
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtStatusCodeP 
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
            Height          =   375
            Left            =   5520
            TabIndex        =   23
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtPremiumCountP 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   22
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtPaymentMethodP 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   21
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtSurrenderValueP 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   20
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtAccruedBonusP 
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
            Height          =   375
            Left            =   5520
            TabIndex        =   19
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label11 
            Caption         =   "Received Todate "
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Suspense Account"
            Height          =   255
            Left            =   3840
            TabIndex        =   33
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Date Of Last Payment"
            Height          =   255
            Left            =   3840
            TabIndex        =   32
            Top             =   780
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "Status Code"
            Height          =   255
            Left            =   3840
            TabIndex        =   31
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Premium Count"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Payment Method"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Surrender Value"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1740
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "Accrued Bonus"
            Height          =   255
            Left            =   3840
            TabIndex        =   27
            Top             =   1740
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame8"
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   8415
         Begin VB.TextBox txtTermOfPolicyP 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   9
            Top             =   660
            Width           =   2055
         End
         Begin VB.TextBox txtMaturityDateP 
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
            Height          =   375
            Left            =   5520
            TabIndex        =   8
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtDateOfCommencementP 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   7
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtProductCodeP 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   6
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtProductP 
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
            Height          =   375
            Left            =   4560
            TabIndex        =   5
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox txtPaymentPeriodP 
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
            Height          =   375
            Left            =   5520
            TabIndex        =   4
            Top             =   660
            Width           =   2535
         End
         Begin VB.TextBox txtExpectedPremium 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   3
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox txtSumAssuredP 
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
            Height          =   375
            Left            =   5520
            TabIndex        =   2
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label Label6 
            Caption         =   "DOC"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Term of Policy"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label Label9 
            Caption         =   "Maturity Date"
            Height          =   255
            Left            =   3840
            TabIndex        =   15
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Product Code"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1140
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Product"
            Height          =   255
            Left            =   3840
            TabIndex        =   13
            Top             =   1140
            Width           =   1695
         End
         Begin VB.Label Label23 
            Caption         =   "Payment Period"
            Height          =   255
            Left            =   3840
            TabIndex        =   12
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label Label15 
            Caption         =   "Premium"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1620
            Width           =   1695
         End
         Begin VB.Label Label24 
            Caption         =   "Sum Assured"
            Height          =   255
            Left            =   3840
            TabIndex        =   10
            Top             =   1620
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmALISMViewPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVIEW As clsALISView


Private Sub Form_Activate()
    Set rsVIEW = New clsALISView
    rsVIEW.loadPolicy
    Set rsVIEW = Nothing
End Sub

