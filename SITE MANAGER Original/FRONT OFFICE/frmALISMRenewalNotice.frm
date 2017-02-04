VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmALISMRenewalNotice 
   Caption         =   "Renewal Notices"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10935
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   270
         Left            =   120
         TabIndex        =   26
         Top             =   6600
         Visible         =   0   'False
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Frame Frame5 
         Height          =   3135
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   9375
         Begin MSComctlLib.ListView ListView2 
            Height          =   2775
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4895
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
            BackColor       =   16761024
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
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   9600
         TabIndex        =   17
         Top             =   3480
         Width           =   1215
         Begin VB.CommandButton cmdDelete 
            Appearance      =   0  'Flat
            Caption         =   "&Re Calc"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdCancel 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Picture         =   "frmALISMRenewalNotice.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdate 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Picture         =   "frmALISMRenewalNotice.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdAddNew 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Picture         =   "frmALISMRenewalNotice.frx":0204
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton cmdEdit 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Picture         =   "frmALISMRenewalNotice.frx":0306
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton cmdSearch 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Picture         =   "frmALISMRenewalNotice.frx":0408
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1560
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   10815
         Begin MSComctlLib.ListView ListView1 
            Height          =   1215
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   2143
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
            BackColor       =   16777215
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
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   10695
         Begin VB.TextBox txtTotalPremium 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtNoOfRecords 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtMonthDescription 
            Appearance      =   0  'Flat
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   720
            Width           =   2040
         End
         Begin VB.TextBox txtPreparedBy 
            Appearance      =   0  'Flat
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
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox txtDatePrepared 
            Appearance      =   0  'Flat
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtRenewalPeriod 
            Appearance      =   0  'Flat
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
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.UpDown UpDownYear 
            Height          =   375
            Left            =   6960
            TabIndex        =   2
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Value           =   2004
            OrigLeft        =   3840
            OrigTop         =   240
            OrigRight       =   4095
            OrigBottom      =   615
            Max             =   2100
            Min             =   2000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownMonth 
            Height          =   375
            Left            =   3240
            TabIndex        =   1
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Value           =   1
            Max             =   12
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtYear 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtMonth 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Description"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Total Premium"
            Height          =   255
            Left            =   4200
            TabIndex        =   30
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "No Of Records"
            Height          =   255
            Left            =   7560
            TabIndex        =   28
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Prepared By"
            Height          =   255
            Left            =   4200
            TabIndex        =   15
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Date Prepared"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Renewal Period"
            Height          =   255
            Left            =   7560
            TabIndex        =   11
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Year"
            Height          =   255
            Left            =   4200
            TabIndex        =   9
            Top             =   300
            Width           =   495
         End
         Begin VB.Label lblReferenceNo 
            Caption         =   "Month"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   300
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmALISMRenewalNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsreceipt As clsALISReceipt

Private Sub cmdAddNew_Click()
        enableALLRECORD
        disableSButtons
End Sub

Private Sub cmdDelete_Click()
        
    If CurrentUserName <> "administrator" Then Exit Sub
    disableButtons
    frmALISMRenewalNotice.cmdUpdate.Enabled = False
    
    Set rsreceipt = New clsALISReceipt
    Set rsreceipt = Nothing
    
    enableButtons
End Sub

Private Sub cmdUpdate_Click()
    Set rsreceipt = New clsALISReceipt
    Set rsreceipt = Nothing
End Sub

Private Sub Form_Activate()
    
    Set rsCONTROL = New ADODB.Recordset
            
    Set rsreceipt = New clsALISReceipt
    Set rsreceipt = Nothing
    enableButtons
    disableALLRECORD
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    frmALISMRenewalNotice.txtRenewalPeriod.Text = Item.Text
    
    Set rsreceipt = New clsALISReceipt
    Set rsreceipt = Nothing

End Sub

Private Sub UpDownMonth_Change()
    With frmALISMRenewalNotice
        
        .txtMonth.Text = .UpDownMonth.Value
        
        If Len(.txtMonth) = 1 Then .txtMonth.Text = "0" + Trim(.txtMonth.Text)
        
        .txtRenewalPeriod.Text = Trim(.txtYear.Text) + "/" + (.txtMonth.Text)
    End With
    
    Set rsreceipt = New clsALISReceipt
    Set rsreceipt = Nothing
        
End Sub

Private Sub UpDownYear_Change()
        With frmALISMRenewalNotice
            .txtYear.Text = .UpDownYear
            .txtRenewalPeriod.Text = Trim(.txtYear.Text) + "/" + (.txtMonth.Text)
        End With
        
        Set rsreceipt = New clsALISReceipt
        Set rsreceipt = Nothing

End Sub
