VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASMTerminateLandLord 
   Caption         =   "Terminate Contract - LandLord"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   9225
   Icon            =   "frmODASMTerminateLandLord.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASMTerminateLandLord.frx":0442
   ScaleHeight     =   5985
   ScaleWidth      =   9225
   Begin VB.Frame Frame12 
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   9015
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   8520
         TabIndex        =   45
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   19726337
         CurrentDate     =   38365
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   5520
         TabIndex        =   44
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   19726337
         CurrentDate     =   38365
      End
      Begin VB.TextBox txtCommencementDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   41
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtExpiryDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         TabIndex        =   40
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5760
         TabIndex        =   39
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtRentRecovered 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   38
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtReoveryRatio 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   36
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtRentPaid 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   33
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtTerminationDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         TabIndex        =   31
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtLandLord 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtNoticeDate 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtPlotNo 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtTerminatedBy 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   24
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cboReasonCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4320
         TabIndex        =   20
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtAgreementDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtLength 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtNarration 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1440
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3120
         Width           =   6255
      End
      Begin VB.TextBox txtContractNo 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtPlotName 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   4335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Terminated Contracts"
         Height          =   1695
         Left            =   120
         TabIndex        =   14
         Top             =   3840
         Width           =   7575
         Begin MSComctlLib.ListView ListView1 
            Height          =   1335
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   2355
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   7800
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
         Begin VB.CommandButton cmdAddNew 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASMTerminateLandLord.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASMTerminateLandLord.frx":0886
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   990
            Width           =   855
         End
         Begin VB.CommandButton cmdSearch 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASMTerminateLandLord.frx":0988
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1365
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASMTerminateLandLord.frx":0A8A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1740
            Width           =   855
         End
         Begin VB.CommandButton cmdCancel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASMTerminateLandLord.frx":0B8C
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2115
            Width           =   855
         End
         Begin VB.CommandButton cmdPrint 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASMTerminateLandLord.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2520
            Width           =   855
         End
      End
      Begin VB.Label Label15 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   43
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Expiry Date"
         Height          =   255
         Left            =   6000
         TabIndex        =   42
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Recovery"
         Height          =   255
         Left            =   5520
         TabIndex        =   37
         Top             =   2670
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Recovery Ratio"
         Height          =   375
         Left            =   3360
         TabIndex        =   35
         Top             =   2610
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Rent Paid"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2670
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Termination Date"
         Height          =   375
         Left            =   6000
         TabIndex        =   32
         Top             =   1650
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Land Lord"
         Height          =   255
         Left            =   6000
         TabIndex        =   30
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Notice Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   27
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Plot No"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Terminated By"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3330
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Plot Name"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Terminated By"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2190
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Date Signed"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Signed By"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   " Reason"
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         Top             =   2190
         Width           =   1455
      End
      Begin VB.Label lblRelationshipCode 
         Caption         =   "Contract No"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   270
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmODASMTerminateLandLord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddNew_Click()
        clearRECORD
        enableRECORD
        disableButtons
End Sub

Private Sub cmdCancel_Click()
        clearALLRECORD
        enableButtons
        disableALLRECORD
End Sub

Private Sub cmdPrint_Click()
        
End Sub

Private Sub cmdUpdate_Click()
        ValidateRECORD
        If bSaveRECORD = True Then
                SaveRECORD
                
                If bSaveRECORD = False Then
                    enableButtons
                    disableALLRECORD
                End If
        End If
End Sub
